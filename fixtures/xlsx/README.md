# XLSX Fixture Corpus

These files are used by the round-trip validation harness (`crates/xlsx-diff`).

Goals:
- Keep fixtures **small** and **in-repo** so CI is fast and deterministic.
- Cover representative XLSX features incrementally (basic cells, formulas, styles, etc.).

## Encrypted / password-to-open workbooks

This corpus is intentionally limited to **ZIP/OPC-based** spreadsheet files that can be
round-tripped by the `crates/xlsx-diff` harness.

This refers to Excel’s **“Encrypt with Password” / “Password to open”** feature (file encryption).
Workbook/worksheet protection (“password to edit”) is **not** encryption and still produces a
normal ZIP/OPC `.xlsx`, so protection-related fixtures can live in this corpus.

Excel “password-to-open” OOXML workbooks (e.g. `.xlsx`, `.xlsm`, `.xlsb`) are **not ZIP archives**
on disk. They are OLE/CFB (Compound File Binary) containers with `EncryptionInfo` and
`EncryptedPackage` streams (MS-OFFCRYPTO). Because the round-trip corpus is collected by
`xlsx-diff::collect_fixture_paths` and then opened as a ZIP archive, encrypted OOXML fixtures must
be **excluded** from this corpus (they would fail ZIP parsing).

Encrypted workbook fixtures live under:

- `fixtures/encrypted/` (see `fixtures/encrypted/README.md`; includes real-world encrypted `.xlsx`/`.xlsb`/`.xls` samples used by end-to-end tests)
- `fixtures/encrypted/ooxml/` (see `fixtures/encrypted/ooxml/README.md`; vendored encrypted `.xlsx`/`.xlsm` corpus including Agile + Standard fixtures, empty-password + Unicode-password samples, macro-enabled `.xlsm` fixtures, plus `*-large.xlsx` multi-segment variants)
- `fixtures/xlsx/encrypted/` (optional/legacy; encrypted OOXML samples co-located with this corpus, but **intentionally skipped** by `xlsx-diff::collect_fixture_paths`)

For background on how Excel encryption works (and how it differs from workbook/worksheet
protection), see `docs/21-encrypted-workbooks.md`.

## Layout

```
fixtures/xlsx/
  basic/
  metadata/
  rich-data/
  richdata/
  images-in-cells/
  formulas/
  styles/
  conditional-formatting/
  hyperlinks/
  charts/
  charts-ex/
  pivots/
  macros/
```

`charts/` and `pivots/` will expand over time as we add more complex corpora.
`macros/` contains small `.xlsm` fixtures used to validate VBA project preservation.
`rich-data/` is reserved for modern Excel 365 “rich value” features (linked data
types and images-in-cells). It is primarily intended for **ground-truth workbooks
saved by Excel**, but it may also contain small **synthetic** fixtures used by
tests (see `rich-data/README.md` and check `docProps/app.xml` for provenance).

## Notable fixtures

- `basic/`:
  - `basic.xlsx` - minimal numeric + inline string.
  - `bool-error.xlsx` - boolean + error cell types.
  - `shared-strings.xlsx` - uses `xl/sharedStrings.xml`.
  - `multi-sheet.xlsx` - 2 worksheets.
  - `comments.xlsx` - legacy comments parts.
  - `grouped-rows.xlsx` - outline/grouped rows metadata.
  - `image.xlsx` - embedded image (`xl/media/image1.png`) + drawing relationship.
  - `rotated-image.xlsx` - **synthetic** rotated image (`xdr:pic` with `<a:xfrm rot="...">`; used for manual verification of transform support).
  - `shape-textbox.xlsx` - **synthetic** DrawingML shape with `<xdr:txBody>` text (used by UI shape text rendering).
  - `rotated-shape.xlsx` - **synthetic** DrawingML shape with `a:xfrm rot="..."` (used for manual verification of rotation/flip support).
  - `image-in-cell.xlsx` - **real Excel** “Place in Cell” image values using `xl/metadata.xml` + `xl/richData/*` (no `xl/cellimages.xml`; see `basic/image-in-cell.md`).
  - `image-in-cell-richdata.xlsx` - **synthetic** minimal in-cell image via rich values (`xl/metadata.xml` + `xl/richData/*`; see `basic/image-in-cell-richdata.md`).
  - `cell-images.xlsx` / `cellimages.xlsx` - **synthetic** in-cell image store part only (`xl/cellImages.xml` or `xl/cellimages.xml` + `.rels`) referencing `xl/media/image1.png` (not Excel ground truth; see `basic/cell-images.md` and `basic/cellimages.md`).
  - `activex-control.xlsx` - minimal ActiveX/form control parts (`xl/ctrlProps/*` + `xl/activeX/*`) with a worksheet `<controls>` fragment.
  - `print-settings.xlsx` - page setup + print titles/areas.
- `formulas/`:
  - `formulas.xlsx` - simple formula + cached result.
  - `formulas-stale-cache.xlsx` - formula with an intentionally stale cached result (used to ensure imports do not auto-recalc).
  - `shared-formula.xlsx` - shared formula range (`t="shared"`) with textless followers.
- `metadata/`:
  - `row-col-properties.xlsx` - custom row height + hidden row, custom column width + hidden column.
  - `data-validation-list.xlsx` - simple list data validation (`<dataValidations>`).
  - `rich-values-vm.xlsx` - **synthetic** workbook-level `xl/metadata.xml` part + worksheet cell `vm="..."` rich-value binding (`futureMetadata` / `rvb`).
  - `defined-names.xlsx` - workbook named ranges (`<definedNames>` in `xl/workbook.xml`).
  - `external-link.xlsx` - minimal external link parts (`xl/externalLinks/externalLink1.xml` + rels).
- `styles/`:
  - `styles.xlsx` - simple bold cell style.
  - `rich-text-shared-strings.xlsx` - shared strings with rich-text runs.
- `rich-data/`:
  - `images-in-cell.xlsx` - **real Excel** images-in-cell fixture including both rich values (`xl/richData/*`) and `xl/cellimages.xml` (see `rich-data/images-in-cell.md`).
  - `richdata-minimal.xlsx` - **synthetic** richData + metadata fixture used by tests (see `rich-data/README.md`).
  - See `rich-data/README.md` for how to generate additional ground-truth Excel fixtures that include
    `xl/metadata.xml`, `xl/richData/*`, and in-sheet `vm`/`cm` cell attributes.
- `images-in-cells/`:
  - `image-in-cell.xlsx` - **Excel** images-in-cells fixture containing both a “Place in Cell” value and an `_xlfn.IMAGE(...)` formula cell.
    Includes `xl/cellimages.xml`, `xl/metadata.xml`, unprefixed `xl/richData/richValue*.xml`, and `xl/media/*` (see `images-in-cells/image-in-cell.md`).
- `richdata/`:
  - `linked-data-types.xlsx` - **synthetic** minimal workbook exercising `xl/metadata.xml` + `xl/richData/*`
    parts and `vm`/`cm` cell attributes (used by `crates/formula-xlsx/tests/linked_data_types_fixture.rs`).
- `conditional-formatting/`:
  - `conditional-formatting.xlsx` - simple `cfRule` example.
  - `conditional-formatting-2007.xlsx` - Excel 2007-style conditional formatting.
  - `conditional-formatting-x14.xlsx` - Excel 2010+ (`x14`) conditional formatting.
- `hyperlinks/`:
  - `hyperlinks.xlsx` - external + internal hyperlinks (tests `hyperlinks` + rels).
- `pivots/`:
  - `pivot-fixture.xlsx` - minimal pivot table parts (cache definition/records + pivotTable).
- `charts/`:
  - `basic-chart.xlsx` - minimal chart parts (drawing + chart referencing sheet data).
  - `rotated-chart.xlsx` - **synthetic** rotated chart frame (`xdr:graphicFrame` with `xdr:xfrm rot="..."`; used for manual verification of chart transform support).
  - `bar.xlsx`, `line.xlsx`, `pie.xlsx`, `scatter.xlsx` - small fixtures for common chart types.
- `charts-ex/`:
  - Parser-focused **ChartEx** fixtures (Excel 2016+ “modern” charts) that include real
    `xl/charts/chartEx1.xml` structures with `cx:*Chart` plot areas and series caches
    (`cx:strCache` / `cx:numCache`). Used by `crates/formula-xlsx/tests/chart_ex_detection.rs`.
- `macros/`:
  - `basic.xlsm` - minimal VBA project preservation fixture.
  - `signed-basic.xlsm` - `basic.xlsm` with a `\x05DigitalSignature` stream containing an
    Office DigSig wrapper (length-prefixed DigSigInfoSerialized-like header) whose PKCS#7/CMS
    `SignedData` payload embeds an Authenticode `SpcIndirectDataContent` digest (self-signed fixture
    cert; no private key).

## Regenerating the initial fixtures

Some `.xlsx` files (the small synthetic ones) are generated without external dependencies:

```bash
python3 fixtures/xlsx/generate_fixtures.py
```

The generator uses deterministic timestamps so diffs are stable.
