# Workstream C: File I/O

> **⛔ STOP. READ [`AGENTS.md`](../AGENTS.md) FIRST. FOLLOW IT COMPLETELY. THIS IS NOT OPTIONAL. ⛔**
>
> This document is supplementary to AGENTS.md. All rules, constraints, and guidelines in AGENTS.md apply to you at all times. Memory limits, build commands, design philosophy—everything.

---

## Mission

Achieve **perfect Excel file compatibility**. Every `.xlsx` file loads perfectly. Every formula works identically. Users can switch with zero friction.

**The goal:** 100% read compatibility, 99.9% calculation fidelity, 97%+ round-trip preservation.

---

## Scope

### Your Crates

| Crate | Purpose |
|-------|---------|
| `crates/formula-xlsx` | XLSX (ECMA-376) reader/writer |
| `crates/formula-xlsb` | XLSB (binary) reader/writer |
| `crates/formula-xls` | Legacy XLS (BIFF) reader |
| `crates/formula-io` | Format detection, streaming I/O |
| `crates/formula-biff` | BIFF record parsing (shared) |
| `crates/formula-vba` | VBA project preservation |

### Your Documentation

- **Primary:** [`docs/02-xlsx-compatibility.md`](../docs/02-xlsx-compatibility.md) — file format handling, preservation strategy
- **Pivots:** [`docs/21-xlsx-pivots.md`](../docs/21-xlsx-pivots.md) — PivotTables/PivotCaches/Slicers/Timelines OpenXML compatibility + roadmap
- **Charts:** [`docs/17-charts.md`](../docs/17-charts.md) — DrawingML chart parsing and round-trip
- **Encrypted workbooks:** [`docs/21-encrypted-workbooks.md`](../docs/21-encrypted-workbooks.md) — password-protected/encrypted Excel files (OOXML `EncryptedPackage`, legacy `.xls` `FILEPASS`)

---

## Key Requirements

### Compatibility Levels

| Level | Description | Target |
|-------|-------------|--------|
| **L1: Read** | File opens, all data visible | 100% |
| **L2: Calculate** | All formulas produce correct results | 99.9% |
| **L3: Render** | Visual appearance matches Excel | 98% |
| **L4: Round-trip** | Save and reopen in Excel with no changes | 97% |
| **L5: Execute** | VBA macros run correctly | 90% (stretch) |

### XLSX Structure

```
workbook.xlsx (ZIP archive)
├── [Content_Types].xml
├── xl/
│   ├── workbook.xml           # Sheet refs, structure
│   ├── styles.xml             # All formatting
│   ├── sharedStrings.xml      # Deduplicated strings
│   ├── calcChain.xml          # Calculation order
│   ├── worksheets/sheet*.xml  # Cell data, formulas
│   ├── drawings/              # Charts, shapes, images
│   ├── charts/                # Chart definitions
│   ├── tables/                # Table definitions
│   ├── pivotTables/           # Pivot tables
│   └── vbaProject.bin         # VBA macros (binary)
└── xl/_rels/                  # Relationships
```

### Critical Preservation Rules

1. **Always store both formula text (`f`) AND cached value (`v`)**
2. **Preserve relationship IDs exactly** — never regenerate
3. **Store `_xlfn.` prefixes** for newer functions
4. **Maintain `calcChain.xml`** for calculation order hints
5. **Use MC namespace** for forward compatibility:

```xml
<mc:AlternateContent>
  <mc:Choice Requires="x14"><!-- Excel 2010+ --></mc:Choice>
  <mc:Fallback><!-- Older apps --></mc:Fallback>
</mc:AlternateContent>
```

### The Five Hardest Problems

1. **Conditional formatting rules** — version divergence between Excel 2007/2010+
2. **Chart fidelity** — DrawingML complexity, ChartEx for newer charts
3. **Date systems** — 1900 vs 1904, Lotus leap year bug (Feb 29, 1900)
4. **Dynamic array function prefixes** — `_xlfn.` handling
5. **VBA macro preservation** — binary vbaProject.bin

### Sheet Metadata

**Tab order:** Order of `<sheet>` elements in `workbook.xml`
**Tab color:** `<tabColor>` in each `worksheet/sheetN.xml`
**Visibility:** `state="hidden"` or `state="veryHidden"`
**Sheet ID:** Stable `sheetId` — never renumber when reordering

---

## Build Commands

```bash
# Build
bash scripts/cargo_agent.sh build --release -p formula-xlsx

# Test
bash scripts/cargo_agent.sh test -p formula-xlsx
bash scripts/cargo_agent.sh test -p formula-xlsb

# Run the OPC-level diff tool (supports .xlsx/.xlsm/.xlsb)
bash scripts/cargo_agent.sh run -p xlsx-diff -- file1.xlsx file2.xlsx
```

Note: `xlsb-diff` is a deprecated compatibility wrapper. Prefer `xlsx-diff` for both `.xlsx` and `.xlsb`.

---

## Performance Targets

| Metric | Target |
|--------|--------|
| File open | <3 seconds for 100MB xlsx |
| Memory | <500MB for 100MB xlsx loaded |
| Round-trip overhead | <5% file size increase |

---

## Test Fixtures

```bash
fixtures/xlsx/           # Test xlsx files
fixtures/charts/         # Chart-specific test files
fixtures/encrypted/      # Password-protected/encrypted workbooks (OLE/CFB wrapper; not part of the ZIP/OPC round-trip corpus)
tools/excel-oracle/      # Excel comparison oracle
tools/corpus/            # Real-world file corpus tools
```

### Round-Trip Testing

```bash
# Use xlsx-diff to compare files (.xlsx/.xlsm/.xlsb)
bash scripts/cargo_agent.sh run -p xlsx-diff -- original.xlsx roundtripped.xlsx

# Run corpus tests
bash scripts/cargo_agent.sh test -p formula-xlsx --test roundtrip
```

---

## Coordination Points

- **Core Engine Team:** Formula parsing, cached values, calculation
- **UI Team:** What they display comes from your parsing
- **Collaboration Team:** CRDT operations map to your data structures

---

## Other Formats

### XLSB (Binary)

- Faster parsing than XLSX (no XML overhead)
- Same structure, different encoding
- Priority for large files

### CSV/Parquet

- Import/export support
- Proper encoding detection (UTF-8, Windows-1252, etc.)
- Parquet for big data workflows

### Legacy XLS (BIFF)

- Read-only support
- Excel 97-2003 format
- Convert to XLSX on save
- `.xls` notes/comments (BIFF NOTE/OBJ/TXO): see [`docs/xls-note-import-hardening.md`](../docs/xls-note-import-hardening.md)

---

## Reference

- ECMA-376 (Office Open XML): https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
- MS-XLSX: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsx/
- MS-XLSB: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/
