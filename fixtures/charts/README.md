# Chart fixture corpus

This directory is the canonical home for chart regression fixtures used by:

- XLSX chart parsing/model extraction tests (`crates/formula-xlsx/tests/chart_fixture_models_match.rs`)
- Fixture completeness + golden PNG assertions (`crates/formula-xlsx/tests/chart_fixture_corpus_complete.rs`)
- Future deterministic chart rendering + visual regression tests (`docs/17-charts.md`)

## Layout

```
fixtures/charts/
  xlsx/                 # Source workbooks (one chart per workbook is preferred)
  models/               # Generated JSON chart models (one JSON file per chart)
  golden/
    excel/              # Golden PNG exports from Excel (fixed pixel size)
```

## Fixture notes

- The classic fixtures (`bar.xlsx`, `line.xlsx`, `pie.xlsx`, `scatter.xlsx`,
  `basic-chart.xlsx`) are mirrored from `fixtures/xlsx/charts/` for backwards
  compatibility with the broader XLSX round-trip corpus.
- The ChartEx-named workbooks (`waterfall.xlsx`, `histogram.xlsx`, etc.) include
  a `xl/charts/chartEx1.xml` part and a `xl/charts/_rels/chart1.xml.rels`
  relationship so we have representative modern-chart OPC graphs in-repo even
  before full ChartEx parsing/rendering lands.
  - Note: these `fixtures/charts/xlsx/*` ChartEx parts are intentionally small
    and may not include the full `cx:*Chart` plot/series structure.
  - For parser development fixtures that **do** include `cx:plotArea`, concrete
    `cx:*Chart` elements, and series caches (`cx:strCache` / `cx:numCache`), see
    `fixtures/xlsx/charts-ex/`.

## Golden image size

Golden PNGs under `fixtures/charts/golden/excel/` are expected to be exported at:

- **800 × 600 px**

### Placeholder goldens
 
Some chart types don't yet have a deterministic Excel-exported golden image
checked into the repo. For these fixtures we commit a generated placeholder PNG
(still **800 × 600 px**) so the corpus remains structurally complete in CI:
 
There are currently two placeholder styles:

1. A “red X” placeholder (light gray background) used by:
   - `area.png`
   - `bar-horizontal.png`
   - `bubble.png`
   - `radar.png`
   - `stock.png`
   - `surface.png`
   - `doughnut.png`
   - `combo-bar-line.png`
   - `map.png`

2. A solid-fill placeholder (single flat color) used by:
   - `bar.png`
   - `bar-gap-overlap.png`
   - `bar-invert-if-negative.png`
   - `basic-chart.png`
   - `box-whisker.png`
   - `funnel.png`
   - `histogram.png`
   - `line.png`
   - `line-smooth.png`
   - `manual-layout.png`
   - `pareto.png`
   - `pie.png`
   - `scatter.png`
   - `sunburst.png`
   - `treemap.png`
   - `waterfall.png`

Replace these with real Excel exports once available (see instructions below).

## Updating / regenerating models

The committed JSON under `fixtures/charts/models/` should always match what the
current Rust parser produces for the corresponding workbook in
`fixtures/charts/xlsx/`.

To regenerate model JSON files:

```bash
# From repo root.
cargo run -p formula-xlsx --bin dump_chart_models -- fixtures/charts/xlsx/bar.xlsx --emit-both-models
```

By default the tool writes files under `fixtures/charts/models/<workbook-stem>/`.
See `--help` for options.

The JSON payload includes the chart index, sheet name, anchor, drawing object
metadata, and one or more parsed `ChartModel`s (see
`formula_model::charts::ChartModel`):

- `drawingRelId`: the `r:id` used in the drawing part to reference the chart
  part (`xl/charts/chartN.xml`).
- `drawingObjectId` / `drawingObjectName`: the DrawingML `<xdr:cNvPr>` id/name
  for the embedded chart object (when present).
- `modelChartSpace`: the classic `c:chartSpace` model parsed from `chart*.xml`.
- `modelChartEx`: optional best-effort `cx:*` (ChartEx) model parsed from
  `chartEx*.xml` when present.

If you only want a single model, `dump_chart_models` supports:

- `--use-chart-object-model` (emits `extract_chart_objects()`'s chosen model,
  preferring ChartEx when present).
- Omitting `--emit-both-models` (legacy schema with a single `model` field,
  parsed from `chart*.xml`).

After regenerating, run:

```bash
cargo test -p formula-xlsx --test chart_fixture_models_match
```

## Fixture completeness checks

CI also runs `cargo test -p formula-xlsx --test chart_fixture_corpus_complete`, which enforces:

- A golden PNG exists for every `fixtures/charts/xlsx/<stem>.xlsx` at
  `fixtures/charts/golden/excel/<stem>.png`.
- Golden PNG dimensions must be exactly **800×600 px**.
- A `fixtures/charts/models/<stem>/` directory exists and contains at least one `chart<N>.json`.

For a subset of fixtures representing ChartEx workbooks, the test also asserts that the package
contains a `xl/charts/chartEx1.xml` part and that `xl/charts/chart1.xml` includes some formatting
elements we care about (e.g. `<c:numFmt>`, `<c:legendPos>`, `<c:dPt>`). If you add a new ChartEx
fixture, update the `chartex_stems` list in
`crates/formula-xlsx/tests/chart_fixture_corpus_complete.rs`.

## Updating / regenerating golden PNGs

1. Open the workbook in desktop Excel (not Google Sheets).
2. Select the chart.
3. Export / copy as picture → **PNG** at **800 × 600 px**.
4. Save to `fixtures/charts/golden/excel/<workbook-stem>.png`.

### Optional: scripted export (Windows + Excel)

If you have Microsoft Excel installed on Windows, you can export all goldens via
COM automation:

```powershell
powershell -ExecutionPolicy Bypass -File tools/excel-oracle/export-chart-goldens.ps1
```

The script exports the first embedded chart in each workbook under
`fixtures/charts/xlsx/` to `fixtures/charts/golden/excel/` at 800×600 px.

The exporter also validates that the PNG dimensions match the expected size, and will warn if the
result appears to be a placeholder/blank image (for example, a near solid-fill output).
