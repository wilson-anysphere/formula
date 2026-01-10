# Charts: DrawingML + ChartEx Compatibility and Rendering

## Overview

Excel charts are defined using **DrawingML** (DrawingML Charts, `c:` namespace) and, for newer chart types, **ChartEx** (a newer schema used by Excel 2016+ for several “modern” charts). Chart support has two non‑negotiable goals:

1. **Round-trip preservation (P0)**: Loading and saving an `.xlsx` must not drop or rewrite chart parts we don’t fully understand.
2. **Rendering fidelity (P1)**: When we *do* render a chart, the result should be visually close to Excel across chart types, formatting, axes, legends, labels, and layout.

This document defines the chart roadmap and a system design that makes it possible to reach broad Excel parity incrementally without data loss.

---

## File/Part Anatomy (What must round-trip)

Charts are not a single XML file. A “chart on a sheet” is a graph of OPC parts:

```
xl/worksheets/sheetN.xml
  └── <drawing r:id="rIdDwg"/>
       └── xl/drawings/drawingN.xml
            └── xdr:twoCellAnchor / oneCellAnchor / absoluteAnchor
                 └── xdr:graphicFrame
                      └── a:graphic/a:graphicData/c:chart r:id="rIdChart"
                           ├── xl/charts/chartM.xml            (DrawingML chartSpace)
                           ├── xl/charts/chartExM.xml          (ChartEx, for some chart types)
                           ├── xl/charts/styleN.xml            (optional preset chart style)
                           ├── xl/charts/colorsN.xml           (optional preset chart color style)
                           └── embedded images (rare, e.g. textures)
```

### Chart-related parts to preserve

At minimum, round-trip must preserve:

- `xl/drawings/drawing*.xml` and its `.rels` exactly (anchors, EMU offsets, relationship ids).
- `xl/charts/chart*.xml` and its `.rels` exactly.
- `xl/charts/chartEx*.xml` and its `.rels` (when present).
- `xl/charts/style*.xml`, `xl/charts/colors*.xml` (when present).
- Any `mc:AlternateContent`, `c:extLst`, `a:extLst` blocks (forward-compat extensions).

**Rule:** if we don’t actively edit charts, the safest default is *byte-for-byte copy* of all chart-related parts on save.

---

## Code Organization (Target Paths)

Charts are an end-to-end feature spanning:

- **XLSX parsing/writing (Rust):** `crates/formula-xlsx/src/drawingml/charts/**`
  - Parse DrawingML charts (`c:chartSpace`) and ChartEx (`cx:chartSpace`) where applicable.
  - Preserve unknown/unsupported chart content for round-trip.
  - Expose a structured `ChartModel` + a raw XML fallback to the application layer.

- **Rendering (Desktop/UI):** `apps/desktop/src/charts/**`
  - Implement a chart scene graph and renderers (Canvas2D first, optionally GPU later).
  - Ensure layout positioning matches drawing anchors (EMUs) and sheet cell geometry.
  - Implement visual regression infrastructure for chart rendering.

These paths are intentionally separated: the core XLSX layer must be *lossless* and deterministic, while the UI renderer can evolve independently as we improve fidelity.

---

## Chart Type Roadmap (Incremental, but comprehensive)

Excel exposes dozens of chart types, but many are parameterizations of a few core geometries. The roadmap below prioritizes what’s required for parity and what’s structurally easiest to implement without sacrificing fidelity.

### Core DrawingML chart types (classic `c:chartSpace`)

| Excel UI | OOXML primary element(s) | Notes |
|---|---|---|
| Column / Bar | `c:barChart` | Includes clustered, stacked, 100% stacked via `c:grouping`. `c:barDir` controls col vs bar. |
| Line | `c:lineChart` | Markers, smoothing, gaps, secondary axis support. |
| Area | `c:areaChart` | Stacked/100% stacked variants exist. |
| Scatter (XY) | `c:scatterChart` | Also the base for many “combo-like” visuals. |
| Pie / Doughnut | `c:pieChart`, `c:doughnutChart` | Exploded slices, leader lines, data labels. |
| Bubble | `c:bubbleChart` | Requires `xVal`, `yVal`, `bubbleSize` series data. |
| Radar | `c:radarChart` | `standard`, `marker`, `filled`. |
| Stock (OHLC) | `c:stockChart` | Typically multiple series (high/low/open/close) + category axis. |
| Combo | multiple chart elements in one `c:plotArea` | Shared axes via `c:axId`. Often secondary axis per subset of series. |

### Modern chart types (often ChartEx, Excel 2016+)

These chart types frequently appear as ChartEx parts (or in extension lists inside a classic chart). The parsing strategy must be robust to both encodings.

| Excel UI | Common storage | Notes |
|---|---|---|
| Waterfall | ChartEx (`cx:*`) | Includes “total” bars, connectors, negative coloring rules. |
| Histogram | ChartEx (`cx:*`) | Requires binning rules (bin width/number of bins, underflow/overflow). |
| Pareto | ChartEx (`cx:*`) | Histogram + cumulative percentage line on secondary axis. |
| Box & Whisker | ChartEx (`cx:*`) | Quartile algorithm, outliers, mean markers, whisker definition. |
| Treemap | ChartEx (`cx:*`) | Hierarchical rectangles, labels by level, color scale. |
| Sunburst | ChartEx (`cx:*`) | Hierarchical rings, label placement. |
| Funnel | ChartEx (`cx:*`) | Stage sizing, gaps, label placement. |
| Map (optional) | ChartEx (`cx:*`) | If unsupported: preserve and show placeholder. |

---

## Parsing Requirements (XLSX → Internal Model → XLSX)

### 1) Lossless ingestion: always keep the original XML

Even when we implement a chart type, we still need the original XML for:

- future fields we don’t model yet,
- Excel extensions (`extLst`) we don’t interpret,
- stable “no-op” round-trips when users only edit cells.

Recommended representation (language-agnostic):

```ts
type OpcPartPath = string;

interface ChartObject {
  // Where it lives on the sheet (from drawing anchors)
  sheetId: string;
  anchor: DrawingAnchor; // twoCell/oneCell/absolute, includes EMU offsets

  // The chart definition parts
  chartPart: { path: OpcPartPath; relsPath: OpcPartPath; xmlBytes: Uint8Array };
  chartExPart?: { path: OpcPartPath; relsPath: OpcPartPath; xmlBytes: Uint8Array };
  stylePart?: { path: OpcPartPath; xmlBytes: Uint8Array };
  colorStylePart?: { path: OpcPartPath; xmlBytes: Uint8Array };

  // Parsed model (optional; only present when we render/edit)
  model?: ChartModel;

  // Optional: structured warnings for unsupported features
  diagnostics: ChartDiagnostic[];
}
```

### 2) Detect and parse both classic and ChartEx forms

Parsing flow:

1. From `sheetN.xml`, follow the drawing relationship to `drawingN.xml`.
2. From each `xdr:*Anchor`, extract:
   - `from/to` cell + EMU offsets (or absolute position),
   - size in EMUs,
   - the target chart relationship id in `graphicData`.
3. Load `xl/charts/chartM.xml`. In addition:
   - If `chartM.xml.rels` includes a relationship to a ChartEx part, load it.
   - Preserve `mc:AlternateContent` and `extLst` blocks exactly.

### 3) Series data resolution (cached vs live)

Excel chart parts often contain both:

- a formula reference to cells (e.g. `Sheet1!$B$2:$B$13`), and
- a cached copy of values (`c:numCache`, `c:strCache`).

**Rendering rule:** prefer cached values unless we can guarantee a correct recalculation of the workbook, including volatile functions and external links. This matches Excel’s own behavior when opening a workbook with stale chart caches.

**Editing rule:** if we ever “rebuild” chart XML, we must decide whether to update caches. For a first implementation, avoid rewriting chart XML unless explicitly editing charts.

---

## Fidelity Requirements (What “acceptable” means)

### Axes

Minimum fidelity features to implement:

- Tick label text and placement (`c:tickLblPos`, `c:tickMark`).
- Number formats (`c:numFmt formatCode="…"`) using the same formatting engine as cells.
- Major/minor gridlines (`c:majorGridlines`, `c:minorGridlines`) including line style.
- Scaling: min/max, log, reverse order (`c:scaling`), crossing rules (`c:crosses`, `c:crossesAt`).
- Axis title (`c:title`) including rich text.

### Legend, titles, data labels

- Chart title: text, rich text, and cell reference title (`c:title` with `c:strRef`).
- Legend: position, overlay behavior, text styling (`c:legend`, `c:legendPos`, `c:overlay`).
- Data labels: series-level defaults and point overrides (`c:dLbls`, `c:dLbl`) for value/category/percent, leader lines, label position.

### Theme + style inheritance (Task 109 dependency)

Charts inherit formatting from:

1. Document theme (`xl/theme/theme1.xml`),
2. Chart preset styles (`xl/charts/style*.xml` + `colors*.xml`),
3. Chart/series/point overrides (`c:spPr`, `c:txPr`, `a:*` drawing properties).

To render consistently, chart color resolution must share the same theme/`schemeClr` machinery as cell fills/fonts.

### Layout + positioning

There are two layouts to get right:

1. **Sheet anchoring** (drawing layer): object position and size in EMUs anchored to cells.
2. **Chart internal layout**: plot area vs legend vs title using `c:layout` / `c:manualLayout`.

#### EMU conversion

- `1 inch = 914400 EMUs`
- `px = (emu / 914400) * dpi` (use a consistent DPI for offscreen rendering; map to device pixels later)

#### Two-cell anchor to pixels

For `xdr:twoCellAnchor`:

```
x = sumPx(colWidth[0..from.col-1]) + emuToPx(from.colOff)
y = sumPx(rowHeight[0..from.row-1]) + emuToPx(from.rowOff)

right  = sumPx(colWidth[0..to.col-1]) + emuToPx(to.colOff)
bottom = sumPx(rowHeight[0..to.row-1]) + emuToPx(to.rowOff)

width  = right - x
height = bottom - y
```

Accurate `colWidthPx` conversion requires Excel’s column width rules (character-based width → pixels). This should be implemented once and reused by images, shapes, and charts.

---

## Rendering Strategy (Consistency across platforms)

We need deterministic rendering across:

- Desktop (Tauri/WebView)
- Web (Canvas/WebGPU)

### Recommended approach: custom renderer via a chart scene graph

Build a platform-agnostic scene graph that resolves chart geometry into:

- paths (rects, lines, arcs),
- fills/strokes,
- text runs (font, size, color),
- clipping regions.

Then implement backends:

- Canvas2D for immediate correctness and broad compatibility,
- optional Skia/WebGPU for performance later.

This avoids the fidelity mismatch you get when mapping Excel charts into third-party chart libraries with different defaults and layout engines.

### Rendering pipeline (high level)

1. Parse chart model (or fallback to placeholder if unsupported).
2. Resolve theme + style inheritance to concrete colors/fonts.
3. Resolve layout (chart area → plot area → axes/legend/title).
4. Project data into screen coordinates (scales).
5. Emit scene graph primitives.
6. Rasterize to the target surface.

**Fallback behavior:** if a chart type is unsupported, render a placeholder rectangle with the chart name/title and keep all parts intact for round-trip.

---

## Testing Plan (fixtures + round-trip + visual regression)

### Fixtures (one workbook per chart family)

Maintain a `fixtures/charts/` corpus created in desktop Excel (not Google Sheets):

- `combo.xlsx`
- `stacked-column.xlsx`, `stacked-bar.xlsx`, `stacked-100.xlsx`
- `bubble.xlsx`
- `radar.xlsx`
- `waterfall.xlsx`
- `histogram.xlsx`, `pareto.xlsx`
- `box-whisker.xlsx`
- `treemap.xlsx`, `sunburst.xlsx`
- `funnel.xlsx`
- `stock-ohlc.xlsx`
- `map.xlsx` (placeholder/preserve)

Each fixture should include:

- non-default theme,
- axis number formats,
- gridlines on/off,
- legend position variations,
- data labels variations,
- at least one series with per-point formatting overrides.

### Round-trip preservation tests

For chart fixtures, implement “no-op” round-trip tests:

1. Load workbook.
2. Save without editing charts.
3. Unzip both and compare:
   - existence of all chart-related parts,
   - relationship ids,
   - byte equality for chart XML parts (preferred) or canonicalized XML equality (acceptable if we reserialize).

### Visual regression tests

For each fixture chart:

- Export a golden PNG from Excel (scripted via VBA or manual export).
- Render in our app at a fixed size and DPI.
- Compare images with a perceptual diff and a strict threshold.

The goal is not perfect pixel-identical output (Excel uses proprietary layout heuristics), but “acceptably close” with stable regressions over time.
