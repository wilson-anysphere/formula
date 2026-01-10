# XLSX Compatibility Layer

## Overview

Perfect XLSX compatibility is the foundation of user trust. Users must be confident that their complex financial models, scientific calculators, and business-critical workbooks will load, calculate, and save without any loss of fidelity.

---

## XLSX File Format Structure

XLSX is a ZIP archive following Open Packaging Conventions (ECMA-376):

```
workbook.xlsx (ZIP archive)
├── [Content_Types].xml          # MIME type declarations
├── _rels/
│   └── .rels                    # Package relationships
├── docProps/
│   ├── app.xml                  # Application properties
│   └── core.xml                 # Core properties (author, dates)
├── xl/
│   ├── workbook.xml             # Workbook structure, sheet refs
│   ├── styles.xml               # All cell formatting
│   ├── sharedStrings.xml        # Deduplicated text strings
│   ├── calcChain.xml            # Calculation order hints
│   ├── theme/
│   │   └── theme1.xml           # Color/font theme
│   ├── worksheets/
│   │   ├── sheet1.xml           # Cell data, formulas
│   │   └── sheet2.xml
│   ├── drawings/
│   │   └── drawing1.xml         # Charts, shapes, images
│   ├── charts/
│   │   └── chart1.xml           # Chart definitions
│   ├── tables/
│   │   └── table1.xml           # Table definitions
│   ├── pivotTables/
│   │   └── pivotTable1.xml      # Pivot table definitions
│   ├── pivotCache/
│   │   ├── pivotCacheDefinition1.xml
│   │   └── pivotCacheRecords1.xml
│   ├── queryTables/
│   │   └── queryTable1.xml      # External data queries
│   ├── connections.xml          # External data connections
│   ├── externalLinks/
│   │   └── externalLink1.xml    # Links to other workbooks
│   ├── customXml/               # Power Query definitions (base64)
│   └── vbaProject.bin           # VBA macros (binary)
└── xl/_rels/
    └── workbook.xml.rels        # Workbook relationships
```

---

## Key Components

### Worksheet XML Structure

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>
  
  <sheetFormatPr defaultRowHeight="15"/>
  
  <cols>
    <col min="1" max="1" width="12.5" style="1" customWidth="1"/>
  </cols>
  
  <sheetData>
    <row r="1" spans="1:3">
      <c r="A1" s="1" t="s">           <!-- t="s" = shared string -->
        <v>0</v>                        <!-- Index into sharedStrings -->
      </c>
      <c r="B1" s="2">                  <!-- No t = number -->
        <v>42.5</v>
      </c>
      <c r="C1" s="3">
        <f>A1+B1</f>                    <!-- Formula -->
        <v>42.5</v>                      <!-- Cached value -->
      </c>
    </row>
  </sheetData>
  
  <conditionalFormatting sqref="A1:C10">
    <cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan">
      <formula>100</formula>
    </cfRule>
  </conditionalFormatting>
  
  <dataValidations count="1">
    <dataValidation type="list" sqref="D1:D100">
      <formula1>"Option1,Option2,Option3"</formula1>
    </dataValidation>
  </dataValidations>
  
  <hyperlinks>
    <hyperlink ref="E1" r:id="rId1"/>
  </hyperlinks>
  
  <mergeCells count="1">
    <mergeCell ref="F1:G2"/>
  </mergeCells>
</worksheet>
```

### Cell Value Types

| Type Attribute | Meaning | Value Content |
|---------------|---------|---------------|
| (absent) | Number | Raw numeric value |
| `t="s"` | Shared String | Index into sharedStrings.xml |
| `t="str"` | Inline String | String in `<v>` element |
| `t="inlineStr"` | Rich Text | `<is><t>text</t></is>` |
| `t="b"` | Boolean | 0 or 1 |
| `t="e"` | Error | Error string (#VALUE!, etc.) |

### Formula Storage

```xml
<!-- Simple formula -->
<c r="A1">
  <f>SUM(B1:B10)</f>
  <v>150</v>
</c>

<!-- Shared formula (for filled ranges) -->
<c r="A1">
  <f t="shared" ref="A1:A10" si="0">B1*2</f>
  <v>10</v>
</c>
<c r="A2">
  <f t="shared" si="0"/>  <!-- References shared formula -->
  <v>20</v>
</c>

<!-- Array formula (legacy CSE style) -->
<c r="A1">
  <f t="array" ref="A1:A5">TRANSPOSE(B1:F1)</f>
  <v>1</v>
</c>

<!-- Dynamic array formula (Excel 365) -->
<c r="A1">
  <f t="array" ref="A1:A5" aca="true">UNIQUE(B1:B100)</f>
  <v>First</v>
</c>

<!-- Formula with _xlfn. prefix for newer functions -->
<c r="A1">
  <f>_xlfn.XLOOKUP(D1,A1:A10,B1:B10)</f>
  <v>Result</v>
</c>
```

### Shared Strings

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
     count="100" uniqueCount="50">
  <si><t>Hello World</t></si>
  <si><t>Another String</t></si>
  <si>
    <r>  <!-- Rich text with formatting runs -->
      <rPr><b/><sz val="12"/></rPr>
      <t>Bold</t>
    </r>
    <r>
      <rPr><sz val="12"/></rPr>
      <t> Normal</t>
    </r>
  </si>
</sst>
```

### Styles

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0.00"/>
  </numFmts>
  
  <fonts count="2">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FF0000FF"/>
      <name val="Arial"/>
    </font>
  </fonts>
  
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  
  <borders count="2">
    <border><!-- empty border --></border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
    </border>
  </borders>
  
  <cellXfs count="3">  <!-- Cell formats reference fonts/fills/borders by index -->
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="164" fontId="1" fillId="1" borderId="1" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>
```

---

## The Five Hardest Compatibility Problems

### 1. Conditional Formatting Version Divergence

Excel 2007 and Excel 2010+ use different XML schemas for the same visual features.

**Excel 2007 Data Bar:**
```xml
<conditionalFormatting sqref="A1:A10">
  <cfRule type="dataBar" priority="1">
    <dataBar>
      <cfvo type="min"/>
      <cfvo type="max"/>
      <color rgb="FF638EC6"/>
    </dataBar>
  </cfRule>
</conditionalFormatting>
```

**Excel 2010+ Data Bar (extended features):**
```xml
<x14:conditionalFormattings xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
    <x14:cfRule type="dataBar" id="{GUID}">
      <x14:dataBar minLength="0" maxLength="100" gradient="0" direction="leftToRight">
        <x14:cfvo type="autoMin"/>
        <x14:cfvo type="autoMax"/>
        <x14:negativeFillColor rgb="FFFF0000"/>
        <x14:axisColor rgb="FF000000"/>
      </x14:dataBar>
    </x14:cfRule>
    <xm:sqref>A1:A10</xm:sqref>
  </x14:conditionalFormatting>
</x14:conditionalFormattings>
```

**Strategy**: 
- Parse both schemas
- Convert internally to unified representation
- Write back preserving original schema version
- Use MC:AlternateContent for cross-version compatibility

### 2. Chart Fidelity (DrawingML)

Charts use DrawingML, a complex XML schema for vector graphics:

```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>
          <c:cat><!-- Categories --></c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$1:$B$10</c:f>
              <c:numCache>
                <c:ptCount val="10"/>
                <c:pt idx="0"><c:v>100</c:v></c:pt>
                <!-- ... -->
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

**Challenges:**
- Different applications render same XML differently
- Absolute positioning in EMUs (English Metric Units)
- Complex inheritance of styles from theme
- Version-specific chart types (Treemap, Sunburst from Excel 2016)

**Strategy:**
- Implement full DrawingML parsing and rendering
- Test extensively against Excel output
- For unsupported chart types, preserve XML and show placeholder

### 3. Date Systems (The Lotus Bug)

Excel supports two date systems:

| System | Epoch | Day 1 |
|--------|-------|-------|
| 1900 (Windows default) | January 1, 1900 | Serial 1 |
| 1904 (Mac legacy) | January 1, 1904 | Serial 0 |

**The Lotus 1-2-3 Bug:**
Excel 1900 system incorrectly treats 1900 as a leap year (it wasn't). February 29, 1900 is serial 60, though this date never existed.

```
Serial 59 = February 28, 1900
Serial 60 = February 29, 1900  ← INVALID DATE
Serial 61 = March 1, 1900
```

**Implications:**
- Dates before March 1, 1900 are off by 1 day
- Mixing 1900 and 1904 workbooks creates 1,462 day differences
- We must emulate this bug for compatibility

**Strategy:**
```typescript
function serialToDate(serial: number, dateSystem: "1900" | "1904"): Date {
  if (dateSystem === "1900") {
    // Emulate Lotus bug
    if (serial < 60) {
      return addDays(new Date(1899, 11, 31), serial);
    } else if (serial === 60) {
      // Invalid date - Feb 29, 1900 didn't exist
      return new Date(1900, 1, 29);  // Represent as-if
    } else {
      // After the bug, off by one
      return addDays(new Date(1899, 11, 30), serial);
    }
  } else {
    return addDays(new Date(1904, 0, 1), serial);
  }
}
```

### 4. Dynamic Array Function Prefixes

Excel 365 introduced dynamic array functions that require `_xlfn.` prefix in file storage:

```xml
<!-- Stored in file -->
<f>_xlfn.UNIQUE(_xlfn.FILTER(A1:A100,B1:B100>0))</f>

<!-- Displayed to user -->
=UNIQUE(FILTER(A1:A100,B1:B100>0))
```

**Functions requiring prefix:**
- UNIQUE, FILTER, SORT, SORTBY, SEQUENCE
- XLOOKUP, XMATCH
- RANDARRAY
- LET, LAMBDA, MAP, REDUCE, SCAN, MAKEARRAY
- Many others added post-2010

**Opening in older Excel:**
```
=_xlfn.XLOOKUP(...)  ← Shown as formula text, #NAME? error
```

**Strategy:**
- Strip prefix on parse for display
- Add prefix on save for file compatibility
- Maintain list of all prefixed functions with version introduced

### 5. VBA Macro Preservation

VBA is stored as binary (`vbaProject.bin`) following OLE compound document format:

```
vbaProject.bin (OLE container)
├── VBA/
│   ├── _VBA_PROJECT     # VBA metadata
│   ├── dir              # Module directory (compressed)
│   ├── Module1          # Module source (compressed)
│   └── ThisWorkbook     # Workbook module
├── PROJECT              # Project properties
└── PROJECTwm            # Project web module
```

**Challenges:**
- Binary format with compression
- Digital signatures must be preserved
- No standard library for creation (only preservation)
- Security implications of execution

**Strategy:**
- Preserve `vbaProject.bin` byte-for-byte on round-trip
- Parse for display/inspection (MS-OVBA specification)
- Defer execution to Phase 2 or via optional component
- Offer migration path to Python/TypeScript

---

## Parsing Libraries and Tools

### By Platform

| Library | Platform | Formulas | Charts | VBA | Pivot |
|---------|----------|----------|--------|-----|-------|
| Open XML SDK | .NET | R/W | Full | Preserve | Partial |
| Apache POI | Java | Eval+R/W | Limited | Preserve | Limited |
| openpyxl | Python | R/W | Good | No | Preserve |
| xlrd/xlwt | Python | Read only | No | No | No |
| SheetJS | JavaScript | Read | Pro only | No | Pro only |
| calamine | Rust | Read | No | No | No |
| rust_xlsxwriter | Rust | Write | Partial | No | No |
| libxlsxwriter | C | Write | Good | No | No |

### Recommended Approach

1. **Reading**: Start with calamine (Rust, fast) for data extraction
2. **Writing**: rust_xlsxwriter for basic files
3. **Full fidelity**: Custom implementation following ECMA-376
4. **Reference**: Apache POI for behavior verification

---

## Round-Trip Preservation Strategy

### Principle: Preserve What We Don't Understand

```typescript
interface XlsxDocument {
  // Fully parsed and modeled
  workbook: Workbook;
  sheets: Sheet[];
  styles: StyleSheet;
  
  // Preserved as raw XML for round-trip
  unknownParts: Map<PartPath, XmlDocument>;
  
  // Preserved byte-for-byte
  binaryParts: Map<PartPath, Uint8Array>;  // vbaProject.bin, etc.
}
```

### Relationship ID Preservation

```xml
<!-- Original -->
<Relationship Id="rId1" Type="...worksheet" Target="worksheets/sheet1.xml"/>

<!-- WRONG: Regenerated IDs break internal references -->
<Relationship Id="rId5" Type="...worksheet" Target="worksheets/sheet1.xml"/>

<!-- CORRECT: Preserve original IDs -->
<Relationship Id="rId1" Type="...worksheet" Target="worksheets/sheet1.xml"/>
```

### Markup Compatibility (MC) Namespace

For forward compatibility with features we don't support:

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice Requires="x14">
    <!-- Excel 2010+ specific content -->
    <x14:sparklineGroups>...</x14:sparklineGroups>
  </mc:Choice>
  <mc:Fallback>
    <!-- Fallback for applications that don't support x14 -->
  </mc:Fallback>
</mc:AlternateContent>
```

**Strategy**: Preserve AlternateContent blocks, process Choice if we support the namespace.

---

## XLSB (Binary Format)

XLSB uses the same ZIP structure but binary records instead of XML:

```
Benefits:
- 2-3x faster to open/save
- 50% smaller file size
- Same feature support as XLSX

Structure:
- Same ZIP layout
- .bin files instead of .xml
- Records: [type: u16][size: u32][data: bytes]
```

### Binary Record Format

```
Record Structure:
┌──────────┬──────────┬────────────────┐
│ Type (2) │ Size (4) │ Data (variable)│
└──────────┴──────────┴────────────────┘

Example - Cell Value Record:
Type: 0x0002 (BrtCellReal)
Size: 8
Data: IEEE 754 double

Example - Formula Record:
Type: 0x0006 (BrtCellFmla)
Size: variable
Data: [value][flags][formula_bytes]
```

**Strategy**: Support XLSB reading for performance, focus XLSX for primary format.

---

## Testing Strategy

### Compatibility Test Suite

1. **Unit tests**: Each file component (cells, formulas, styles, etc.)
2. **Integration tests**: Complex workbooks with multiple features
3. **Round-trip tests**: Load → Save → Load, compare
4. **Cross-application tests**: Save from us, open in Excel; save from Excel, open in us
5. **Real-world corpus**: Test against collection of user-submitted files

### Test File Categories

| Category | Examples | Focus Areas |
|----------|----------|-------------|
| Basic | Simple data, formulas | Core functionality |
| Styling | Rich formatting, themes | Visual fidelity |
| Charts | All chart types | DrawingML rendering |
| Pivots | Complex pivot tables | Pivot cache, definitions |
| External | Links, queries | Connection handling |
| Large | 1M+ rows | Performance, memory |
| Legacy | Excel 97-2003 | .xls conversion |
| Complex | Financial models | Everything together |

### Automated Comparison

```typescript
interface ComparisonResult {
  identical: boolean;
  differences: Difference[];
}

interface Difference {
  path: string;  // e.g., "xl/worksheets/sheet1.xml/row[5]/c[3]/v"
  type: "missing" | "added" | "changed";
  original?: string;
  modified?: string;
  severity: "critical" | "warning" | "info";
}

async function compareWorkbooks(
  original: XlsxFile,
  roundTripped: XlsxFile
): Promise<ComparisonResult> {
  // Compare structure
  // Compare XML content with normalization
  // Compare binary parts byte-for-byte
  // Report all differences with severity
}
```

### Implemented Round-Trip Harness (xlsx-diff)

This repository includes a small XLSX fixture corpus and a part-level diff tool to
validate load → save → diff round-trips:

- Fixtures live under `fixtures/xlsx/**` (kept intentionally small).
- The diff tool is implemented in Rust: `crates/xlsx-diff`.

Run a diff locally:

```bash
cargo run -p xlsx-diff --bin xlsx_diff -- original.xlsx roundtripped.xlsx
```

Run the fixture harness (used by CI):

```bash
cargo test -p xlsx-diff --test roundtrip_fixtures
```

Current normalization rules (to reduce false positives):

- Ignore whitespace-only XML text nodes unless `xml:space="preserve"` is set.
- Sort XML attributes (namespace declarations are ignored; resolved URIs are used instead).
- Sort `<Relationships>` entries by `(Id, Type, Target)`.
- Sort `[Content_Types].xml` entries by `(Default.Extension | Override.PartName)`.
- Sort worksheet `<sheetData>` rows/cells by their `r` attributes.

Current severity policy (subject to refinement as the writer matures):

- **critical**: missing parts, changes in `[Content_Types].xml`, changes in `*.rels`, any binary part diffs.
- **warning**: non-essential parts like themes / calcChain, extra parts.
- **info**: metadata-only changes under `docProps/*`.

---

## Performance Considerations

### Streaming Parsing

For large files, don't load entire XML into memory:

```typescript
// BAD: Load entire file
const xml = await parseXml(await readFile(path));
const cells = xml.querySelectorAll("c");

// GOOD: Stream parsing
const parser = new SaxParser();
parser.on("element:c", (cell) => {
  processCell(cell);
});
await parser.parseStream(fileStream);
```

### Lazy Loading

Don't parse everything upfront:

```typescript
class LazyWorksheet {
  private parsed = false;
  private xmlPath: string;
  private data?: SheetData;
  
  async getData(): Promise<SheetData> {
    if (!this.parsed) {
      this.data = await this.parse();
      this.parsed = true;
    }
    return this.data!;
  }
}
```

### Parallel Processing

Parse independent parts concurrently:

```typescript
const [workbook, styles, sharedStrings] = await Promise.all([
  parseWorkbook(archive),
  parseStyles(archive),
  parseSharedStrings(archive),
]);

// Then parse sheets (which depend on above)
const sheets = await Promise.all(
  workbook.sheets.map(s => parseSheet(archive, s, styles, sharedStrings))
);
```

---

## Future Considerations

1. **Excel for Web compatibility**: Some features differ in web version
2. **Google Sheets export**: Import/export from Google's format
3. **Numbers compatibility**: Apple's format for Mac users
4. **OpenDocument (ODS)**: LibreOffice compatibility
5. **New Excel features**: Monitor Excel updates for new XML schemas
