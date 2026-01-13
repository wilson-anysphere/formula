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


def _zip_write_bytes(zf: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(name, date_time=EPOCH)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.create_system = 0
    zf.writestr(info, data)


def write_xlsx(
    path: pathlib.Path,
    sheet_xmls: list[str],
    styles_xml: str,
    *,
    sheet_names: list[str] | None = None,
    shared_strings_xml: str | None = None,
    workbook_xml_override: str | None = None,
    workbook_rels_extra: list[str] | None = None,
    content_types_extra_overrides: list[str] | None = None,
    extra_parts: dict[str, str | bytes] | None = None,
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
                extra_overrides=content_types_extra_overrides,
            ),
        )
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(
            zf,
            "xl/workbook.xml",
            workbook_xml_override if workbook_xml_override is not None else workbook_xml(sheet_names),
        )
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(
                sheet_count=len(sheet_xmls),
                include_shared_strings=shared_strings_xml is not None,
                extra_relationships=workbook_rels_extra,
            ),
        )
        for idx, sheet_xml in enumerate(sheet_xmls, start=1):
            _zip_write(zf, f"xl/worksheets/sheet{idx}.xml", sheet_xml)
        _zip_write(zf, "xl/styles.xml", styles_xml)
        if shared_strings_xml is not None:
            _zip_write(zf, "xl/sharedStrings.xml", shared_strings_xml)
        if extra_parts is not None:
            for name, data in extra_parts.items():
                if isinstance(data, bytes):
                    _zip_write_bytes(zf, name, data)
                else:
                    _zip_write(zf, name, data)


def content_types_xml(
    *, sheet_count: int, include_shared_strings: bool, extra_overrides: list[str] | None = None
) -> str:
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
    if extra_overrides:
        overrides.extend(extra_overrides)
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


def workbook_xml_with_extra(sheet_names: list[str], extra: str) -> str:
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
%s
</workbook>
"""
        % ("\n".join(sheets), extra)
    )


def workbook_rels_xml(
    *,
    sheet_count: int,
    include_shared_strings: bool,
    extra_relationships: list[str] | None = None,
) -> str:
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

    if extra_relationships:
        rels.extend(extra_relationships)

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


def styles_varied_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  <fonts count="4">
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
    <font>
      <i/>
      <u/>
      <sz val="14"/>
      <color rgb="FFFF0000"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
    <font>
      <strike/>
      <sz val="11"/>
      <color rgb="FF00AA00"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFFF00"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FF00B0F0"/>
        <bgColor indexed="64"/>
      </patternFill>
    </fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF000000"/></left>
      <right style="thin"><color rgb="FF000000"/></right>
      <top style="thin"><color rgb="FF000000"/></top>
      <bottom style="thin"><color rgb="FF000000"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="11">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="7" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
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


def sheet_formulas_stale_cache_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1">
        <f>A1+B1</f>
        <v>999</v>
      </c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_shared_formula_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1">
        <f t="shared" ref="A1:A3" si="0">B1*2</f>
        <v>2</v>
      </c>
      <c r="B1"><v>1</v></c>
    </row>
    <row r="2">
      <c r="A2">
        <f t="shared" si="0"/>
        <v>4</v>
      </c>
      <c r="B2"><v>2</v></c>
    </row>
    <row r="3">
      <c r="A3">
        <f t="shared" si="0"/>
        <v>6</v>
      </c>
      <c r="B3"><v>3</v></c>
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


def sheet_bool_error_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="b"><v>1</v></c>
      <c r="B1" t="e"><v>#DIV/0!</v></c>
      <c r="C1" t="e"><v>#FIELD!</v></c>
      <c r="D1" t="e"><v>#CONNECT!</v></c>
      <c r="E1" t="e"><v>#BLOCKED!</v></c>
      <c r="F1" t="e"><v>#UNKNOWN!</v></c>
    </row>
    <row r="2">
      <c r="A2" t="b"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_extended_errors_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e"><v>#GETTING_DATA</v></c>
      <c r="B1" t="e"><v>#FIELD!</v></c>
      <c r="C1" t="e"><v>#CONNECT!</v></c>
      <c r="D1" t="e"><v>#BLOCKED!</v></c>
      <c r="E1" t="e"><v>#UNKNOWN!</v></c>
      <c r="F1" t="e"><v>#DIV/0!</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_varied_styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1" t="inlineStr"><is><t>Bold</t></is></c>
      <c r="B1" s="2" t="inlineStr"><is><t>Italic+Underline</t></is></c>
      <c r="C1" s="3" t="inlineStr"><is><t>Strike</t></is></c>
      <c r="D1" s="4" t="inlineStr"><is><t>Fill</t></is></c>
      <c r="E1" s="5" t="inlineStr"><is><t>Border</t></is></c>
      <c r="F1" s="6" t="inlineStr"><is><t>Center Wrap</t></is></c>
      <c r="G1" s="7"><v>0.25</v></c>
      <c r="H1" s="8"><v>1234.5</v></c>
      <c r="I1" s="9"><v>44927</v></c>
      <c r="J1" s="10"><v>42.5</v></c>
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


def sheet_date_iso_cell_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="d"><v>2024-01-01T00:00:00Z</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_rich_values_vm_xml() -> str:
    # Minimal sheet with a `vm="1"` attribute on a cell. This is used by newer Excel
    # rich-value features (ex: images-in-cell) to bind a cell to a record in
    # `xl/metadata.xml`.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def rich_values_metadata_xml() -> str:
    # This is a deliberately tiny `xl/metadata.xml` that includes:
    # - `<metadataTypes>` (type table)
    # - `<futureMetadata>` carrying a rich-value binding (`rvb`)
    # - `<valueMetadata>` used by worksheet cell `vm="..."` attributes
    #
    # The specific schema is future-facing; we only need this part to exist and be
    # preserved byte-for-byte during round-trip edits.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" maxSupportedVersion="120000"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{3E2803F5-59A4-4A43-8C86-93BA0C219F4F}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"""


def sheet_richdata_minimal_xml() -> str:
    # Minimal sheet with `vm`/`cm` attributes (Excel uses these to bind a cell/value to records
    # in `xl/metadata.xml`). The RichData tables themselves live under `xl/richData/*`.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="1"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def metadata_richdata_rels_xml() -> str:
    # Relationships from `xl/metadata.xml` to the workbook-global RichData tables.
    # Relationship Type URIs are Excel-version dependent; for preservation tests they only need
    # to be well-formed and stable.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2017/relationships/richValueTypes" Target="richData/richValueTypes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2017/relationships/richValueStructure" Target="richData/richValueStructure.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2017/relationships/richValueRel" Target="richData/richValueRel.xml"/>
  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/relationships/richValue" Target="richData/richValue.xml"/>
</Relationships>
"""


def richdata_rich_value_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypes xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <types>
    <type id="0" name="com.microsoft.excel.image" structure="s_image"/>
  </types>
</rvTypes>
"""


def richdata_rich_value_structure_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStruct xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <structures>
    <structure id="s_image">
      <member name="imageRel" kind="rel"/>
      <member name="altText" kind="string"/>
    </structure>
  </structures>
</rvStruct>
"""


def richdata_rich_value_rel_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rels>
    <rel r:id="rId1"/>
  </rels>
</rvRel>
"""


def richdata_rich_value_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <values>
    <!-- Relationship index 0 => richValueRel.xml entry 0 => r:id => image1.png -->
    <rv type="0">
      <v kind="rel">0</v>
      <v kind="string">Alt text</v>
    </rv>
  </values>
</rvData>
"""


def richdata_rich_value_rel_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"""


def sheet_row_col_properties_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="3" max="3" hidden="1"/>
    <col min="2" max="2" width="25" customWidth="1"/>
  </cols>
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1" t="inlineStr"><is><t>Wide</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Hidden</t></is></c>
    </row>
    <row r="2" ht="30" customHeight="1">
      <c r="A2"><v>2</v></c>
    </row>
    <row r="3" hidden="1">
      <c r="A3"><v>3</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_data_validation_list_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Pick</t></is></c>
    </row>
  </sheetData>
  <dataValidations count="1">
    <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
      <formula1>"Yes,No"</formula1>
    </dataValidation>
  </dataValidations>
</worksheet>
"""


def sheet_external_link_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1">
        <f>'[external.xlsx]Sheet1'!A1</f>
        <v>0</v>
      </c>
    </row>
  </sheetData>
</worksheet>
"""


def external_link1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <externalBook r:id="rId1">
    <sheetNames>
      <sheetName val="Sheet1"/>
    </sheetNames>
  </externalBook>
</externalLink>
"""


def external_link1_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="external.xlsx" TargetMode="External"/>
</Relationships>
"""


def sheet_hyperlinks_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>External</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Internal</t></is></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Email</t></is></c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1" display="Example" tooltip="Go to example"/>
    <hyperlink ref="A2" location="Sheet2!B2" display="Jump"/>
    <hyperlink ref="A3" r:id="rId2"/>
  </hyperlinks>
</worksheet>
"""


def sheet_hyperlinks_target_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2" t="inlineStr"><is><t>Target</t></is></c>
    </row>
  </sheetData>
</worksheet>
"""


def sheet_hyperlinks_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="mailto:test@example.com" TargetMode="External"/>
</Relationships>
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


def drawing_rotated_chart_xml() -> str:
    # Variant of `drawing1_xml()` with a rotated chart frame (`xdr:graphicFrame`).
    #
    # Rotation is expressed in DrawingML's 60000ths-of-a-degree units.
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
        <xdr:cNvPr id="2" name="Rotated Chart 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm rot="2700000">
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


def sheet_smartart_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>SmartArt</t></is></c>
    </row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>
"""


def drawing_smartart_xml() -> str:
    # Minimal DrawingML structure that references a SmartArt/diagram payload via
    # `dgm:relIds` relationship IDs. The diagram parts live under `xl/diagrams/*`
    # and are referenced from `xl/drawings/_rels/drawing1.xml.rels`.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
          xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>6</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>10</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="SmartArt 1"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""


def drawing_smartart_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
</Relationships>
"""


def diagram_data1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
"""


def diagram_layout1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
"""


def diagram_quick_style1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
"""


def diagram_colors1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>
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


def write_rotated_chart_xlsx(path: pathlib.Path) -> None:
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
        _zip_write(zf, "xl/drawings/drawing1.xml", drawing_rotated_chart_xml())
        _zip_write(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing1_rels_xml())
        _zip_write(zf, "xl/charts/chart1.xml", chart1_xml())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def write_hyperlinks_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1", "Sheet2"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(
            zf,
            "[Content_Types].xml",
            content_types_xml(sheet_count=2, include_shared_strings=False),
        )
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=2, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_hyperlinks_xml())
        _zip_write(zf, "xl/worksheets/sheet2.xml", sheet_hyperlinks_target_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet_hyperlinks_rels_xml())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def content_types_image_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def sheet_image_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Image</t></is></c>
    </row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>
"""


def sheet1_image_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"""


def sheet_rotated_shape_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Rotated shape</t></is></c>
    </row>
  </sheetData>
  <drawing r:id="rId1"/>
</worksheet>
"""


def sheet1_rotated_shape_rels_xml() -> str:
    return sheet1_image_rels_xml()


def drawing_rotated_shape_xml() -> str:
    # Minimal drawing with a rotated shape (`xdr:sp`).
    #
    # Rotation is expressed in DrawingML's 60000ths-of-a-degree units.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>2</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="1828800" cy="914400"/>
    <xdr:sp>
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="Rotated Shape 1"/>
        <xdr:cNvSpPr/>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm rot="2700000">
          <a:off x="0" y="0"/>
          <a:ext cx="1828800" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>
"""


def drawing_empty_rels_xml() -> str:
    # Drawing parts always have a `.rels` sidecar, even if empty.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>
"""


def drawing_image_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>4</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>6</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Picture 1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""


def drawing_rotated_image_xml() -> str:
    # Minimal drawing with a rotated image (`xdr:pic`).
    #
    # Rotation is expressed in DrawingML's 60000ths-of-a-degree units.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>4</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>6</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="Rotated Picture 1"/>
        <xdr:cNvPicPr/>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:xfrm rot="5400000">
          <a:off x="0" y="0"/>
          <a:ext cx="1828800" cy="914400"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""


def drawing_image_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"""


def one_by_one_png_bytes() -> bytes:
    # A minimal 1x1 PNG (transparent). Keeping this inline avoids extra binary
    # assets in the generator script.
    import base64

    return base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMBAe5Z5E8AAAAASUVORK5CYII="
    )


def content_types_cellimages_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def cellimages_xml() -> str:
    # Synthetic "in-cell image store" (`xl/cellimages.xml`) fixture.
    #
    # Real Excel workbooks can include a `cellimages` store part, but Excel's
    # observed schema shape is richer (e.g. it can embed an `xdr:pic` subtree).
    # This minimal variant exists as an on-disk fixture for manual debugging and
    # unknown-part preservation tests.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>
"""


def cellimages_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"""


def write_cellimages_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_cellimages_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(
                sheet_count=1,
                include_shared_strings=False,
                extra_relationships=[
                    '  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2022/relationships/cellImages" Target="cellimages.xml"/>'
                ],
            ),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_basic_xml())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())
        _zip_write(zf, "xl/cellimages.xml", cellimages_xml())
        _zip_write(zf, "xl/_rels/cellimages.xml.rels", cellimages_rels_xml())
        _zip_write_bytes(zf, "xl/media/image1.png", one_by_one_png_bytes())


def content_types_image_in_cell_richdata_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def sheet_image_in_cell_richdata_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="0"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>
"""


def metadata_image_in_cell_richdata_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000" maxSupportedVersion="120000"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="1">
    <bk>
      <extLst>
        <ext uri="{3E2803F5-59A4-4A43-8C86-93BA0C219F4F}">
          <xlrd:rvb i="0"/>
        </ext>
      </extLst>
    </bk>
  </futureMetadata>
  <valueMetadata count="1">
    <bk>
      <rc t="1" v="0"/>
    </bk>
  </valueMetadata>
</metadata>
"""


def rich_value_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv s="0" t="image">
    <v kind="rel">0</v>
  </rv>
</rvData>
"""


def rich_value_rel_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRel xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <rel r:id="rId1"/>
</richValueRel>
"""


def rich_value_rel_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"""


def workbook_rels_image_in_cell_richdata_extra() -> str:
    return '  <Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>'


def write_image_in_cell_richdata_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_image_in_cell_richdata_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(
                sheet_count=1,
                include_shared_strings=False,
                extra_relationships=[
                    workbook_rels_image_in_cell_richdata_extra(),
                    '  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValue" Target="richData/richValue.xml"/>',
                    '  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/richValueRel" Target="richData/richValueRel.xml"/>',
                ],
            ),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_image_in_cell_richdata_xml())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())
        _zip_write(zf, "xl/metadata.xml", metadata_image_in_cell_richdata_xml())
        _zip_write(zf, "xl/richData/richValue.xml", rich_value_xml())
        _zip_write(zf, "xl/richData/richValueRel.xml", rich_value_rel_xml())
        _zip_write(
            zf,
            "xl/richData/_rels/richValueRel.xml.rels",
            rich_value_rel_rels_xml(),
        )
        _zip_write_bytes(zf, "xl/media/image1.png", one_by_one_png_bytes())


def content_types_background_image_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def sheet_background_image_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
    </row>
  </sheetData>
  <picture r:id="rId1"/>
</worksheet>
"""


def sheet1_background_image_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"""


def write_background_image_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_background_image_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=1, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_background_image_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet1_background_image_rels_xml())
        _zip_write_bytes(zf, "xl/media/image1.png", one_by_one_png_bytes())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def content_types_ole_object_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def sheet_ole_object_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>OLE</t></is></c>
    </row>
  </sheetData>
  <oleObjects>
    <oleObject progId="Package" dvAspect="DVASPECT_ICON" oleUpdate="OLEUPDATE_ALWAYS" shapeId="1" r:id="rId2"/>
  </oleObjects>
</worksheet>
"""


def sheet1_ole_object_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>
"""


def write_ole_object_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_ole_object_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=1, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_ole_object_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet1_ole_object_rels_xml())
        _zip_write_bytes(zf, "xl/embeddings/oleObject1.bin", b"OLE\x00OBJECT\x01")
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def content_types_chart_sheet_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def workbook_rels_chart_sheet_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def sheet_chart_sheet_data_only_xml() -> str:
    # Data referenced by xl/charts/chart1.xml.
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
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
</worksheet>
"""


def chartsheet1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <drawing r:id="rId1"/>
</chartsheet>
"""


def chartsheet1_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>
"""


def write_chart_sheet_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1", "Chart1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_chart_sheet_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(zf, "xl/_rels/workbook.xml.rels", workbook_rels_chart_sheet_xml())
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_chart_sheet_data_only_xml())
        _zip_write(zf, "xl/chartsheets/sheet1.xml", chartsheet1_xml())
        _zip_write(zf, "xl/chartsheets/_rels/sheet1.xml.rels", chartsheet1_rels_xml())
        _zip_write(zf, "xl/drawings/drawing1.xml", drawing1_xml())
        _zip_write(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing1_rels_xml())
        _zip_write(zf, "xl/charts/chart1.xml", chart1_xml())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def write_image_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_image_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=1, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_image_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet1_image_rels_xml())
        _zip_write(zf, "xl/drawings/drawing1.xml", drawing_image_xml())
        _zip_write(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing_image_rels_xml())

        _zip_write_bytes(zf, "xl/media/image1.png", one_by_one_png_bytes())

        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def write_rotated_image_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_image_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=1, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_image_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet1_image_rels_xml())
        _zip_write(zf, "xl/drawings/drawing1.xml", drawing_rotated_image_xml())
        _zip_write(zf, "xl/drawings/_rels/drawing1.xml.rels", drawing_image_rels_xml())

        _zip_write_bytes(zf, "xl/media/image1.png", one_by_one_png_bytes())

        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def content_types_activex_control_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.ctrlProps+xml"/>
  <Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


def sheet_activex_control_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <controls>
    <control r:id="rId1" name="Control 1" shapeId="1025"/>
  </controls>
</worksheet>
"""


def sheet1_control_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../ctrlProps/ctrlProp1.xml"/>
</Relationships>
"""


def ctrl_prop1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ctrlProp xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>
"""


def ctrl_prop1_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControl" Target="../activeX/activeX1.xml"/>
</Relationships>
"""


def active_x1_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<activeX xmlns="http://schemas.microsoft.com/office/2006/activeX"/>
"""


def active_x1_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>
</Relationships>
"""


def active_x1_bin_bytes() -> bytes:
    return b"FORMULA-ACTIVEX"


def write_activex_control_xlsx(path: pathlib.Path) -> None:
    sheet_names = ["Sheet1"]
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        path.unlink()

    with zipfile.ZipFile(path, "w") as zf:
        _zip_write(zf, "[Content_Types].xml", content_types_activex_control_xml())
        _zip_write(zf, "_rels/.rels", package_rels_xml())
        _zip_write(zf, "docProps/core.xml", core_props_xml())
        _zip_write(zf, "docProps/app.xml", app_props_xml(sheet_names))
        _zip_write(zf, "xl/workbook.xml", workbook_xml(sheet_names))
        _zip_write(
            zf,
            "xl/_rels/workbook.xml.rels",
            workbook_rels_xml(sheet_count=1, include_shared_strings=False),
        )
        _zip_write(zf, "xl/worksheets/sheet1.xml", sheet_activex_control_xml())
        _zip_write(zf, "xl/worksheets/_rels/sheet1.xml.rels", sheet1_control_rels_xml())
        _zip_write(zf, "xl/ctrlProps/ctrlProp1.xml", ctrl_prop1_xml())
        _zip_write(zf, "xl/ctrlProps/_rels/ctrlProp1.xml.rels", ctrl_prop1_rels_xml())
        _zip_write(zf, "xl/activeX/activeX1.xml", active_x1_xml())
        _zip_write(zf, "xl/activeX/_rels/activeX1.xml.rels", active_x1_rels_xml())
        _zip_write_bytes(zf, "xl/activeX/activeX1.bin", active_x1_bin_bytes())
        _zip_write(zf, "xl/styles.xml", styles_minimal_xml())


def main() -> None:
    write_xlsx(
        ROOT / "basic" / "basic.xlsx",
        [sheet_basic_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "basic" / "bool-error.xlsx",
        [sheet_bool_error_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "basic" / "extended-errors.xlsx",
        [sheet_extended_errors_xml()],
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
        ROOT / "formulas" / "formulas-stale-cache.xlsx",
        [sheet_formulas_stale_cache_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "formulas" / "shared-formula.xlsx",
        [sheet_shared_formula_xml()],
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
    write_xlsx(
        ROOT / "styles" / "varied_styles.xlsx",
        [sheet_varied_styles_xml()],
        styles_varied_xml(),
    )
    write_chart_xlsx(ROOT / "charts" / "basic-chart.xlsx")
    write_rotated_chart_xlsx(ROOT / "charts" / "rotated-chart.xlsx")
    write_image_xlsx(ROOT / "basic" / "image.xlsx")
    write_rotated_image_xlsx(ROOT / "basic" / "rotated-image.xlsx")
    write_xlsx(
        ROOT / "basic" / "rotated-shape.xlsx",
        [sheet_rotated_shape_xml()],
        styles_minimal_xml(),
        content_types_extra_overrides=[
            '  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
        ],
        extra_parts={
            "xl/worksheets/_rels/sheet1.xml.rels": sheet1_rotated_shape_rels_xml(),
            "xl/drawings/drawing1.xml": drawing_rotated_shape_xml(),
            "xl/drawings/_rels/drawing1.xml.rels": drawing_empty_rels_xml(),
        },
        sheet_names=["Sheet1"],
    )
    write_image_in_cell_richdata_xlsx(ROOT / "basic" / "image-in-cell-richdata.xlsx")
    write_cellimages_xlsx(ROOT / "basic" / "cellimages.xlsx")
    write_background_image_xlsx(ROOT / "basic" / "background-image.xlsx")
    write_ole_object_xlsx(ROOT / "basic" / "ole-object.xlsx")
    write_chart_sheet_xlsx(ROOT / "charts" / "chart-sheet.xlsx")
    write_hyperlinks_xlsx(ROOT / "hyperlinks" / "hyperlinks.xlsx")
    write_activex_control_xlsx(ROOT / "basic" / "activex-control.xlsx")
    write_xlsx(
        ROOT / "basic" / "smartart.xlsx",
        [sheet_smartart_xml()],
        styles_minimal_xml(),
        content_types_extra_overrides=[
            '  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>',
            '  <Override PartName="/xl/diagrams/data1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/>',
            '  <Override PartName="/xl/diagrams/layout1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/>',
            '  <Override PartName="/xl/diagrams/quickStyle1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/>',
            '  <Override PartName="/xl/diagrams/colors1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/>',
        ],
        extra_parts={
            "xl/worksheets/_rels/sheet1.xml.rels": sheet1_drawing_rels_xml(),
            "xl/drawings/drawing1.xml": drawing_smartart_xml(),
            "xl/drawings/_rels/drawing1.xml.rels": drawing_smartart_rels_xml(),
            "xl/diagrams/data1.xml": diagram_data1_xml(),
            "xl/diagrams/layout1.xml": diagram_layout1_xml(),
            "xl/diagrams/quickStyle1.xml": diagram_quick_style1_xml(),
            "xl/diagrams/colors1.xml": diagram_colors1_xml(),
        },
        sheet_names=["Sheet1"],
    )

    write_xlsx(
        ROOT / "metadata" / "date-iso-cell.xlsx",
        [sheet_date_iso_cell_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "metadata" / "row-col-properties.xlsx",
        [sheet_row_col_properties_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "metadata" / "data-validation-list.xlsx",
        [sheet_data_validation_list_xml()],
        styles_minimal_xml(),
    )
    write_xlsx(
        ROOT / "metadata" / "rich-values-vm.xlsx",
        [sheet_rich_values_vm_xml()],
        styles_minimal_xml(),
        workbook_xml_override=workbook_xml_with_extra(
            ["Sheet1"],
            """  <metadata r:id="rId3"/>""",
        ),
        workbook_rels_extra=[
            '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>'
        ],
        content_types_extra_overrides=[
            '  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>'
        ],
        extra_parts={
            "xl/metadata.xml": rich_values_metadata_xml(),
        },
        sheet_names=["Sheet1"],
    )
    write_xlsx(
        ROOT / "rich-data" / "richdata-minimal.xlsx",
        [sheet_richdata_minimal_xml()],
        styles_minimal_xml(),
        workbook_xml_override=workbook_xml_with_extra(
            ["Sheet1"],
            """  <metadata r:id="rId3"/>""",
        ),
        workbook_rels_extra=[
            '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>'
        ],
        content_types_extra_overrides=[
            '  <Default Extension="png" ContentType="image/png"/>',
            '  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>',
            '  <Override PartName="/xl/richData/richValue.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>',
            '  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>',
            '  <Override PartName="/xl/richData/richValueTypes.xml" ContentType="application/vnd.ms-excel.richvaluetypes+xml"/>',
            '  <Override PartName="/xl/richData/richValueStructure.xml" ContentType="application/vnd.ms-excel.richvaluestructure+xml"/>',
        ],
        extra_parts={
            "xl/metadata.xml": rich_values_metadata_xml(),
            "xl/_rels/metadata.xml.rels": metadata_richdata_rels_xml(),
            "xl/richData/richValue.xml": richdata_rich_value_xml(),
            "xl/richData/richValueRel.xml": richdata_rich_value_rel_xml(),
            "xl/richData/richValueTypes.xml": richdata_rich_value_types_xml(),
            "xl/richData/richValueStructure.xml": richdata_rich_value_structure_xml(),
            "xl/richData/_rels/richValueRel.xml.rels": richdata_rich_value_rel_rels_xml(),
            "xl/media/image1.png": one_by_one_png_bytes(),
        },
        sheet_names=["Sheet1"],
    )
    write_xlsx(
        ROOT / "metadata" / "defined-names.xlsx",
        [sheet_basic_xml()],
        styles_minimal_xml(),
        workbook_xml_override=workbook_xml_with_extra(
            ["Sheet1"],
            """  <definedNames>
    <definedName name="ZedName">Sheet1!$B$1</definedName>
    <definedName name="MyRange">Sheet1!$A$1:$A$1</definedName>
    <definedName name="ErrName">#N/A</definedName>
  </definedNames>""",
        ),
        sheet_names=["Sheet1"],
    )
    write_xlsx(
        ROOT / "metadata" / "external-link.xlsx",
        [sheet_external_link_xml()],
        styles_minimal_xml(),
        workbook_xml_override=workbook_xml_with_extra(
            ["Sheet1"],
            """  <externalReferences>
    <externalReference r:id="rId3"/>
  </externalReferences>""",
        ),
        workbook_rels_extra=[
            '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
        ],
        content_types_extra_overrides=[
            '  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>'
        ],
        extra_parts={
            "xl/externalLinks/externalLink1.xml": external_link1_xml(),
            "xl/externalLinks/_rels/externalLink1.xml.rels": external_link1_rels_xml(),
        },
        sheet_names=["Sheet1"],
    )

    # Directory scaffold for future corpora (kept empty for now).
    for name in ["charts", "pivots", "macros"]:
        (ROOT / name).mkdir(parents=True, exist_ok=True)

    print("Generated XLSX fixtures under", ROOT)


if __name__ == "__main__":
    main()
