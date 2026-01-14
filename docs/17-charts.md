# Chart System Design

## Overview

Excel's charting system is built on DrawingML (ECMA-376), one of the most complex parts of the file format. Perfect chart compatibility requires understanding both the file format representation and Excel's rendering behavior.

---

## Current Implementation (Rust)

This repo currently has **read + preserve** support for embedded charts in `formula-xlsx`, plus a **best-effort parser** into `formula-model` types. Rendering is not implemented yet.

### Extraction: `XlsxPackage::extract_chart_objects()`

Implemented in `crates/formula-xlsx/src/workbook.rs`.

At a high level the extractor:

1. Walks drawing parts (`xl/drawings/drawing*.xml`) and finds embedded chart references inside
   `<xdr:graphicFrame> … <c:chart r:id="…"/> … </xdr:graphicFrame>`.
2. Extracts and returns:
   - The chart anchor (`formula_model::drawings::Anchor`) derived from `<xdr:twoCellAnchor>`,
     `<xdr:oneCellAnchor>`, or `<xdr:absoluteAnchor>`.
   - The raw `<xdr:graphicFrame>` subtree as a string (`drawing_frame_xml`) for debugging and
     future round-trip editing.
3. Resolves the drawing relationship id (`r:id`) to an `xl/charts/chartN.xml` part.
4. Reads `xl/charts/_rels/chartN.xml.rels` to discover optional related parts:
    - **ChartEx part** (`xl/charts/chartExN.xml`) when relationship type/target contains `chartEx`
      (Excel 2016+ “modern” charts).
    - **Chart style** (`xl/charts/styleN.xml`) and **chart colors** (`xl/charts/colorsN.xml`) via
      relationship type or filename heuristics.
    - **Chart user shapes** (`xl/drawings/drawingN.xml`) via the
      `…/relationships/chartUserShapes` relationship type (callouts, overlays, etc).
    - The `.rels` XML bytes are stored alongside extracted parts so that callers can follow
      relationships without retaining the full `XlsxPackage` in memory.
5. Returns a `formula_xlsx::drawingml::charts::ChartObject` containing:
    - `parts.chart` (the classic `c:chartSpace` part) plus optional `parts.chart_ex`, `parts.style`,
      `parts.colors`, and `parts.user_shapes` as raw bytes (`OpcPart`).
   - `model: Option<ChartModel>` parsed best-effort (see below).
   - `diagnostics` for missing parts / parsing failures while extracting.

> Note: `XlsxPackage::extract_charts()` exists as an older/simple API that returns a smaller
> `formula_model::charts::Chart` shape. New work should prefer `extract_chart_objects()`.

### Parsing: `drawingml::charts::parse_chart_space`

Implemented in `crates/formula-xlsx/src/drawingml/charts/parse_chart_space.rs`.

`parse_chart_space()` parses a classic DrawingML chart (`<c:chartSpace>`) into
`formula_model::charts::ChartModel` with best-effort coverage:

- Detects chart kind by selecting a **primary** `*Chart` element under `c:plotArea` (the first
  supported chart-type element, falling back to the first element when all are unknown).
- Supports **combo charts** where `<c:plotArea>` contains multiple `*Chart` elements by producing
  `PlotAreaModel::Combo` and tagging each `SeriesModel` with `plot_index` to indicate which subplot
  it belongs to.
- Parses:
  - chart title (`c:title`) and legend (`c:legend`) including basic text styling (`c:txPr`)
  - plot-area chart settings (e.g. `barDir`, `grouping`, `varyColors`, `scatterStyle`)
  - axes for `c:catAx` and `c:valAx` (id, position, scaling, number format, tick label position,
      major gridline presence + some inline styling)
  - series identity/order metadata (`c:ser/c:idx/@val`, `c:ser/c:order/@val`)
  - series formulas (`tx`, `cat`, `val`, `xVal`, `yVal`) and cached values (`strCache`, `numCache`)
  - series data label settings (`c:ser/c:dLbls`) including `showVal`, `showCatName`, `showSerName`,
    `dLblPos`, and `numFmt`
  - some inline formatting: `spPr` shape styles, markers, and per-point `c:dPt` overrides
- Records non-fatal parse limitations as `ChartModel.diagnostics` warnings (for example:
  `mc:AlternateContent` (branch selection is heuristic), `c:extLst`, unsupported chart/axis types).

### Parsing: `drawingml::charts::parse_chart_ex`

Implemented in `crates/formula-xlsx/src/drawingml/charts/parse_chart_ex.rs`.

ChartEx (`<cx:chartSpace>`) is used for modern chart types like histogram, waterfall, treemap, etc.
The current parser intentionally produces a **placeholder** `ChartModel`:

- `chart_kind` is always `ChartKind::Unknown { name: "ChartEx:<kind>" }` (best-effort inferred).
- Series formulas + cached values are extracted when present.
- Series identity/order metadata is extracted best-effort when present (e.g. `cx:ser/cx:idx/@val`,
  `cx:ser/cx:order/@val`).
- Title and legend are extracted best-effort when present.
- Axes, styles, and chart-type-specific semantics are not modeled yet; diagnostics note the
  placeholder status.

`extract_chart_objects()` currently prefers `parse_chart_ex()` when a ChartEx part is present.

### Supported chart kinds today

From `parse_chart_space()`:

- **Area** (`ChartKind::Area`): `<c:areaChart>` and `<c:area3DChart>`.
- **Bar/Column** (`ChartKind::Bar`): `<c:barChart>` and `<c:bar3DChart>`.
  - `barDir="col"` vs `barDir="bar"` distinguishes column vs bar.
- **Bubble** (`ChartKind::Bubble`): `<c:bubbleChart>`.
- **Doughnut** (`ChartKind::Doughnut`): `<c:doughnutChart>`.
- **Line** (`ChartKind::Line`): `<c:lineChart>` and `<c:line3DChart>`.
- **Pie** (`ChartKind::Pie`): `<c:pieChart>` and `<c:pie3DChart>`.
- **Radar** (`ChartKind::Radar`): `<c:radarChart>`.
- **Scatter** (`ChartKind::Scatter`): `<c:scatterChart>`.
- **Stock** (`ChartKind::Stock`): `<c:stockChart>`.
- **Surface** (`ChartKind::Surface`): `<c:surfaceChart>` and `<c:surface3DChart>`.
- **Combo** (`PlotAreaModel::Combo`): multiple chart-type elements overlaid within one plot area
  (e.g. barChart + lineChart).
- Everything else becomes `ChartKind::Unknown { name: "<elementName>" }` (series/axes may still be
  extracted).

From `parse_chart_ex()`:

- Everything is currently `ChartKind::Unknown { name: "ChartEx:<kind>" }`.

### Lossless preservation vs parsed model

**Preserved losslessly (byte-for-byte at the OPC part payload level):**

- Chart OPC parts referenced by the drawing:
  - `xl/charts/chartN.xml` (`ChartParts.chart.bytes`)
    - and its `.rels` payload (`ChartParts.chart.rels_bytes`) when present
  - `xl/charts/chartExN.xml` + `xl/charts/_rels/chartExN.xml.rels` when present
    - (available as `ChartParts.chart_ex.bytes` / `ChartParts.chart_ex.rels_bytes`)
  - `xl/charts/styleN.xml` / `xl/charts/colorsN.xml` when present
  - `xl/drawings/drawingN.xml` + `xl/drawings/_rels/drawingN.xml.rels` when present (chart user shapes)
- Raw `<xdr:graphicFrame>` XML (`ChartObject.drawing_frame_xml`) is extracted exactly as a slice
  of the drawing part.

**Parsed into `formula_model::charts::ChartModel` (best-effort):**

- A subset of chart metadata (kind, title, legend, axes, series formulas + caches, some inline
  formatting).

Anything not modeled in `ChartModel` today is still preserved in the original XML bytes, but will
not be available to consumers operating purely on the parsed model until the roadmap items below
are implemented.

---

## Known Gaps (Roadmap)

This list is intentionally written as a work checklist and should map 1:1 to planned parser/model
work in `formula-xlsx` + `formula-model`.

- [ ] **Combo charts**: support plot areas with multiple `*Chart` elements (mixed chart types,
      secondary axes, mixed stacking/grouping). (**chartSpace parsing is implemented**; remaining
      work is rendering semantics / secondary-axis behavior).
- [x] **More classic chart types**: area, radar, bubble, stock, surface, doughnut (modeled as
      first-class `ChartKind`/`PlotAreaModel` variants).
- [ ] **ChartEx modeling**: parse `cx:chartSpace` into a first-class model (histogram, waterfall,
      treemap, sunburst, funnel, box & whisker, Pareto, …) instead of placeholder `Unknown`.
- [ ] **ChartStyle / ChartColorStyle parts**: parse and apply `xl/charts/styleN.xml` and
      `xl/charts/colorsN.xml` (currently extracted but not interpreted).
- [ ] **`mc:AlternateContent` handling**: fully honor `mc:Choice/@Requires` (current parsing flattens
      AlternateContent with heuristic Choice/Fallback selection).
- [ ] **`c:extLst` handling**: model important extensions (today we only record a warning).
- [ ] **Literal series data**: support `c:strLit` / `c:numLit` (today we only handle `strRef` /
      `numRef`).
- [ ] **Multi-level categories**: support `c:multiLvlStrRef` / `c:multiLvlStrCache` and related
      hierarchical category structures.
- [ ] **Axis titles**: parse `c:*Ax/c:title` (currently ignored).
- [x] **Series data labels**: parse `c:ser/c:dLbls` (show value/category/series name, position,
      number format).
- [ ] **Per-point data label overrides**: parse `c:dLbls/c:dLbl` (per-point overrides, rich text).

---

## Fixtures & Tests

The chart regression corpus lives under `fixtures/charts/` (see also
`fixtures/charts/README.md` for the canonical instructions).

For parser development of **ChartEx** specifically (series caches, kind detection, etc.), see the
smaller “real ChartEx” fixtures under `fixtures/xlsx/charts-ex/` (each includes a `chartEx` part
with concrete `cx:*Chart` elements and cached series points).

### Adding a new fixture

1. Create a workbook with (ideally) **one embedded chart**.
2. Save it to `fixtures/charts/xlsx/<stem>.xlsx`.
3. Export a golden image from Excel at **800×600 px** to
   `fixtures/charts/golden/excel/<stem>.png`.
4. Generate model JSON files:

   ```bash
   cargo run -p formula-xlsx --bin dump_chart_models -- fixtures/charts/xlsx/<stem>.xlsx --emit-both-models
   ```

   This writes `fixtures/charts/models/<stem>/chart<N>.json` (one JSON per extracted chart),
   including the drawing relationship/object metadata (`drawingRelId`, `drawingObjectId`,
   `drawingObjectName`) plus the parsed chart models (`modelChartSpace`, optional `modelChartEx`).
5. Commit the XLSX, the model JSON(s), and the golden PNG.

### What the tests enforce

The Rust test suite treats this corpus as **source-of-truth** for chart parsing regression.

- `crates/formula-xlsx/tests/chart_fixture_corpus_complete.rs`
  - Every `fixtures/charts/xlsx/*.xlsx` must have a corresponding golden PNG.
  - Golden images must be exactly **800×600 px**.
  - Every fixture must have a `fixtures/charts/models/<stem>/` directory with at least one
    `chart<N>.json`.
  - Certain ChartEx fixtures are additionally checked for presence of `chartEx1.xml` and some
    representative formatting/features inside `chart1.xml`.
- `crates/formula-xlsx/tests/chart_fixture_models_match.rs`
  - Loads each fixture with `XlsxPackage::extract_chart_objects()`.
  - Requires an exact 1:1 match between extracted chart count and the number of `chart<N>.json`
    files under `fixtures/charts/models/<stem>/`.
  - Asserts the parsed `ChartModel`s match the committed JSON:
    - `modelChartSpace`: `parse_chart_space(chart1.xml)`
    - `modelChartEx`: optional `parse_chart_ex(chartEx1.xml)` when present.

## Chart Types to Support

### Priority 0 (Launch Required)

| Type | DrawingML Element | Complexity |
|------|-------------------|------------|
| Column | `<c:barChart>` with `barDir="col"` | Medium |
| Bar | `<c:barChart>` with `barDir="bar"` | Medium |
| Line | `<c:lineChart>` | Medium |
| Pie | `<c:pieChart>` | Medium |
| Area | `<c:areaChart>` | Medium |
| Scatter (XY) | `<c:scatterChart>` | Medium |

### Priority 1 (Power Users)

| Type | DrawingML Element | Complexity |
|------|-------------------|------------|
| Combo | Multiple chart types overlaid | High |
| Doughnut | `<c:doughnutChart>` | Low |
| Radar | `<c:radarChart>` | Medium |
| Stock (OHLC) | `<c:stockChart>` | High |
| Surface | `<c:surfaceChart>` | High |
| Bubble | `<c:bubbleChart>` | Medium |

### Priority 2 (Specialty)

| Type | DrawingML Element | Complexity |
|------|-------------------|------------|
| Treemap | `<cx:chart>` (ChartEx) | Very High |
| Sunburst | `<cx:chart>` (ChartEx) | Very High |
| Waterfall | `<cx:chart>` (ChartEx) | Very High |
| Funnel | `<cx:chart>` (ChartEx) | High |
| Box & Whisker | `<cx:chart>` (ChartEx) | High |
| Histogram | `<cx:chart>` (ChartEx) | Medium |
| Map | Geographic data | Very High |

---

## File Format Structure

### Chart Part Location

```
xl/
├── workbook.xml
├── worksheets/
│   └── sheet1.xml          <- Contains <drawing> reference
├── drawings/
│   ├── drawing1.xml        <- Anchors chart to cell range
│   └── _rels/
│       └── drawing1.xml.rels  <- Links to chart parts
└── charts/
    ├── chart1.xml          <- Classic chartSpace chart definition
    ├── chartEx1.xml        <- Optional ChartEx part (modern chart types)
    ├── colors1.xml         <- Optional chart color style
    ├── style1.xml          <- Optional chart style
    └── _rels/
        ├── chart1.xml.rels   <- Links to ChartEx/style/colors/external data
        └── chartEx1.xml.rels <- Optional rels for ChartEx resources
```

### DrawingML Chart XML Structure

```xml
<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:p>
            <a:r>
              <a:t>Sales by Region</a:t>
            </a:r>
          </a:p>
        </c:rich>
      </c:tx>
    </c:title>
    
    <c:plotArea>
      <c:layout/>  <!-- Positioning -->
      
      <!-- Chart type specific element -->
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        
        <c:ser>  <!-- Series 1 -->
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx>
            <c:strRef>
              <c:f>Sheet1!$B$1</c:f>  <!-- Series name -->
            </c:strRef>
          </c:tx>
          <c:cat>  <!-- Categories (X-axis) -->
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>  <!-- Values (Y-axis) -->
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
        
        <c:axId val="1"/>  <!-- Category axis ID -->
        <c:axId val="2"/>  <!-- Value axis ID -->
      </c:barChart>
      
      <c:catAx>  <!-- Category axis definition -->
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>  <!-- bottom -->
        <c:crossAx val="2"/>
      </c:catAx>
      
      <c:valAx>  <!-- Value axis definition -->
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>  <!-- left -->
        <c:crossAx val="1"/>
        <c:numFmt formatCode="General"/>
      </c:valAx>
      
    </c:plotArea>
    
    <c:legend>
      <c:legendPos val="r"/>  <!-- right -->
    </c:legend>
    
  </c:chart>
  
  <!-- Cached data for offline viewing -->
  <c:externalData r:id="rId1">
    <c:autoUpdate val="0"/>
  </c:externalData>
  
</c:chartSpace>
```

### ChartEx (Excel 2016+ Charts)

Treemaps, sunbursts, waterfalls use the newer ChartEx format.

In OPC terms this typically shows up as a separate `xl/charts/chartExN.xml` part referenced from
`xl/charts/_rels/chartN.xml.rels` via a `…/relationships/chartEx` relationship.

```xml
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="treemap">
          <cx:dataLabels/>
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>
  <cx:chartData>
    <cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$10</cx:f>
      </cx:strDim>
      <cx:numDim type="size">
        <cx:f>Sheet1!$B$2:$B$10</cx:f>
      </cx:numDim>
    </cx:data>
  </cx:chartData>
</cx:chartSpace>
```

---

## Implementation Architecture

### Chart Data Model

#### Current Rust model (`formula-model`)

Today, chart parsing/extraction in Rust is centered around:

- `formula_xlsx::XlsxPackage::extract_chart_objects()` → `formula_xlsx::drawingml::charts::ChartObject`
- `formula_xlsx::drawingml::charts::{parse_chart_space, parse_chart_ex}` → `formula_model::charts::ChartModel`

Simplified shape:

```rust
use formula_model::charts::ChartModel;
use formula_model::drawings::Anchor;
use formula_xlsx::drawingml::charts::{ChartParts, ChartDiagnostic};

pub struct ChartObject {
    pub sheet_name: Option<String>,
    pub drawing_part: String,
    pub anchor: Anchor,
    pub drawing_frame_xml: String,
    pub parts: ChartParts,           // raw OPC parts (chart/chartEx/style/colors/userShapes) + `.rels` bytes
    pub model: Option<ChartModel>,   // parsed best-effort (may be None)
    pub diagnostics: Vec<ChartDiagnostic>,
}
```

#### Future UI model (TypeScript)

```typescript
interface Chart {
  id: string;
  type: ChartType;
  title?: ChartTitle;
  plotArea: PlotArea;
  legend?: Legend;
  series: ChartSeries[];
  axes: ChartAxis[];
  
  // Positioning
  anchor: ChartAnchor;
  
  // Styling
  style?: ChartStyle;
  colorScheme?: ColorScheme;
}

interface ChartSeries {
  index: number;
  name?: CellReference | string;
  categories?: CellReference;
  values: CellReference;
  
  // Visual properties
  fill?: Fill;
  line?: LineProperties;
  marker?: MarkerProperties;
  dataLabels?: DataLabelProperties;
  
  // Cached data (for round-trip)
  cachedCategories?: (string | number)[];
  cachedValues?: number[];
}

interface ChartAxis {
  id: number;
  type: 'category' | 'value' | 'date' | 'series';
  position: 'top' | 'bottom' | 'left' | 'right';
  title?: AxisTitle;
  scaling: AxisScaling;
  numberFormat?: string;
  tickMarks?: TickMarkProperties;
  gridLines?: GridLineProperties;
  crossesAt?: number | 'autoZero' | 'max' | 'min';
}
```

### Rendering Pipeline

```
┌──────────────────┐
│   Chart Model    │
│  (from parser)   │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  Layout Engine   │
│  - Calculate     │
│    plot area     │
│  - Position axes │
│  - Scale data    │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Render Pipeline  │
│  1. Background   │
│  2. Grid lines   │
│  3. Plot area    │
│  4. Data series  │
│  5. Axes         │
│  6. Labels       │
│  7. Legend       │
│  8. Title        │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  Canvas Output   │
└──────────────────┘
```

### Canvas Rendering

```typescript
class ChartRenderer {
  private ctx: CanvasRenderingContext2D;
  private chart: Chart;
  private bounds: Rectangle;
  
  render(): void {
    this.calculateLayout();
    this.renderBackground();
    this.renderGridLines();
    this.renderPlotArea();
    this.renderSeries();
    this.renderAxes();
    this.renderLegend();
    this.renderTitle();
  }
  
  private renderSeries(): void {
    for (const series of this.chart.series) {
      switch (this.chart.type) {
        case 'bar':
        case 'column':
          this.renderBarSeries(series);
          break;
        case 'line':
          this.renderLineSeries(series);
          break;
        case 'pie':
          this.renderPieSeries(series);
          break;
        case 'scatter':
          this.renderScatterSeries(series);
          break;
        // ... more types
      }
    }
  }
  
  private renderBarSeries(series: ChartSeries): void {
    const { ctx } = this;
    const data = this.getSeriesData(series);
    const barWidth = this.calculateBarWidth();
    
    for (let i = 0; i < data.length; i++) {
      const x = this.dataToCanvasX(i);
      const y = this.dataToCanvasY(data[i].value);
      const height = this.plotArea.bottom - y;
      
      ctx.fillStyle = this.getSeriesColor(series, i);
      ctx.fillRect(x, y, barWidth, height);
      
      if (series.line) {
        ctx.strokeStyle = series.line.color;
        ctx.lineWidth = series.line.width;
        ctx.strokeRect(x, y, barWidth, height);
      }
    }
  }
}
```

---

## Data Binding

### Live Updates

Charts must update when source data changes:

```typescript
class ChartDataBinding {
  private chart: Chart;
  private worksheet: Worksheet;
  private subscriptions: Subscription[] = [];
  
  bind(): void {
    for (const series of this.chart.series) {
      // Watch category range
      if (series.categories) {
        this.subscriptions.push(
          this.worksheet.watchRange(series.categories, () => {
            this.refreshSeriesCategories(series);
          })
        );
      }
      
      // Watch value range
      this.subscriptions.push(
        this.worksheet.watchRange(series.values, () => {
          this.refreshSeriesValues(series);
        })
      );
      
      // Watch series name cell
      if (series.name && typeof series.name !== 'string') {
        this.subscriptions.push(
          this.worksheet.watchCell(series.name, () => {
            this.refreshSeriesName(series);
          })
        );
      }
    }
  }
  
  unbind(): void {
    this.subscriptions.forEach(s => s.unsubscribe());
    this.subscriptions = [];
  }
}
```

### Caching for Round-Trip

Excel stores cached data in charts for offline viewing. We must preserve this:

```xml
<c:val>
  <c:numRef>
    <c:f>Sheet1!$B$2:$B$5</c:f>
    <c:numCache>
      <c:formatCode>General</c:formatCode>
      <c:ptCount val="4"/>
      <c:pt idx="0"><c:v>100</c:v></c:pt>
      <c:pt idx="1"><c:v>200</c:v></c:pt>
      <c:pt idx="2"><c:v>150</c:v></c:pt>
      <c:pt idx="3"><c:v>175</c:v></c:pt>
    </c:numCache>
  </c:numRef>
</c:val>
```

---

## Sparklines

Sparklines are mini-charts inside cells:

```xml
<x14:sparklineGroups xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
  <x14:sparklineGroup type="line" displayEmptyCellsAs="gap">
    <x14:colorSeries theme="4"/>
    <x14:colorNegative theme="5"/>
    <x14:colorMarkers theme="4" tint="-0.499984740745262"/>
    <x14:sparklines>
      <x14:sparkline>
        <xm:f>Sheet1!A1:L1</xm:f>  <!-- Data range -->
        <xm:sqref>M1</xm:sqref>    <!-- Display cell -->
      </x14:sparkline>
    </x14:sparklines>
  </x14:sparklineGroup>
</x14:sparklineGroups>
```

### Sparkline Types

| Type | Description |
|------|-------------|
| `line` | Line chart (most common) |
| `column` | Bar chart |
| `stacked` | Win/loss visualization |

---

## Interactivity

### Chart Selection

```typescript
class ChartInteraction {
  private chart: Chart;
  private selectedElement: ChartElement | null = null;
  
  handleClick(x: number, y: number): void {
    const element = this.hitTest(x, y);
    
    if (element) {
      this.select(element);
      this.emit('select', element);
    } else {
      this.deselect();
    }
  }
  
  private hitTest(x: number, y: number): ChartElement | null {
    // Test in reverse render order (top elements first)
    
    // Test data points
    for (const series of this.chart.series) {
      const point = this.hitTestSeries(series, x, y);
      if (point) return { type: 'dataPoint', series, point };
    }
    
    // Test legend items
    const legendItem = this.hitTestLegend(x, y);
    if (legendItem) return { type: 'legendItem', ...legendItem };
    
    // Test axes
    const axis = this.hitTestAxis(x, y);
    if (axis) return { type: 'axis', axis };
    
    // Test title
    if (this.hitTestTitle(x, y)) {
      return { type: 'title' };
    }
    
    // Test chart area (for moving/resizing)
    if (this.hitTestChartArea(x, y)) {
      return { type: 'chart' };
    }
    
    return null;
  }
}
```

### Tooltips

```typescript
class ChartTooltip {
  show(element: ChartElement, x: number, y: number): void {
    const content = this.formatTooltip(element);
    
    this.tooltipEl.innerHTML = content;
    this.tooltipEl.style.left = `${x + 10}px`;
    this.tooltipEl.style.top = `${y + 10}px`;
    this.tooltipEl.style.display = 'block';
  }
  
  private formatTooltip(element: ChartElement): string {
    switch (element.type) {
      case 'dataPoint':
        return `
          <div class="chart-tooltip">
            <div class="series-name">${element.series.name}</div>
            <div class="category">${element.point.category}</div>
            <div class="value">${this.formatValue(element.point.value)}</div>
          </div>
        `;
      // ... other element types
    }
  }
}
```

---

## Performance Considerations

### Large Datasets

For charts with thousands of points:

```typescript
class OptimizedChartRenderer {
  // Downsample for display
  private downsample(data: DataPoint[], maxPoints: number): DataPoint[] {
    if (data.length <= maxPoints) return data;
    
    // LTTB (Largest Triangle Three Buckets) algorithm
    return lttbDownsample(data, maxPoints);
  }
  
  // Use WebGL for very large datasets
  private shouldUseWebGL(pointCount: number): boolean {
    return pointCount > 10000;
  }
  
  // Batch canvas operations
  private renderOptimized(points: DataPoint[]): void {
    const ctx = this.ctx;
    
    // Single path for all line segments
    ctx.beginPath();
    ctx.moveTo(points[0].x, points[0].y);
    
    for (let i = 1; i < points.length; i++) {
      ctx.lineTo(points[i].x, points[i].y);
    }
    
    ctx.stroke();  // Single draw call
  }
}
```

### Animation

```typescript
class ChartAnimator {
  animate(from: ChartState, to: ChartState, duration: number): void {
    const start = performance.now();
    
    const tick = (now: number) => {
      const elapsed = now - start;
      const t = Math.min(elapsed / duration, 1);
      const eased = this.easeOutCubic(t);
      
      const state = this.interpolate(from, to, eased);
      this.render(state);
      
      if (t < 1) {
        requestAnimationFrame(tick);
      }
    };
    
    requestAnimationFrame(tick);
  }
  
  private interpolate(from: ChartState, to: ChartState, t: number): ChartState {
    // Interpolate each data point
    return {
      series: from.series.map((s, i) => ({
        ...s,
        values: s.values.map((v, j) => 
          v + (to.series[i].values[j] - v) * t
        )
      }))
    };
  }
}
```

---

## Testing Strategy

Current parser regression tests are implemented in Rust and driven by the
`fixtures/charts/` corpus (see **Fixtures & Tests** above). The TypeScript examples below describe
future renderer-level visual regression and round-trip testing.

### Visual Regression Tests

```typescript
describe('Chart Rendering', () => {
  it('should match Excel column chart', async () => {
    const xlsx = await loadTestFile('charts/column-basic.xlsx');
    const chart = xlsx.sheets[0].charts[0];
    
    const canvas = renderChart(chart, 800, 600);
    const screenshot = canvas.toDataURL();
    
    expect(screenshot).toMatchImageSnapshot({
      failureThreshold: 0.01,  // 1% pixel difference allowed
      failureThresholdType: 'percent'
    });
  });
});
```

### Round-Trip Tests

```typescript
describe('Chart Preservation', () => {
  it('should preserve chart on round-trip', async () => {
    const original = await loadTestFile('charts/complex.xlsx');
    const saved = await saveAndReload(original);
    
    // Compare chart XML
    expect(saved.charts[0].xml).toEqualXML(original.charts[0].xml);
    
    // Verify in Excel
    // (automated Excel comparison in CI)
  });
});
```

---

## AI Integration

### Natural Language Chart Creation

```typescript
// AI tool for creating charts
const createChartTool = {
  name: 'create_chart',
  description: 'Create a chart from data',
  parameters: {
    type: { type: 'string', enum: ['bar', 'line', 'pie', 'scatter', 'area'] },
    dataRange: { type: 'string', description: 'Range like A1:D10' },
    title: { type: 'string', optional: true },
    xAxisLabel: { type: 'string', optional: true },
    yAxisLabel: { type: 'string', optional: true }
  }
};

// Usage: "Create a bar chart of sales by region from columns A and B"
```

### Chart Suggestions

```typescript
class ChartSuggestionEngine {
  suggest(data: CellRange): ChartSuggestion[] {
    const analysis = this.analyzeData(data);
    const suggestions: ChartSuggestion[] = [];
    
    // Time series → Line chart
    if (analysis.hasTimeColumn && analysis.hasNumericColumns) {
      suggestions.push({
        type: 'line',
        reason: 'Data appears to be a time series',
        confidence: 0.9
      });
    }
    
    // Categories + values → Bar chart
    if (analysis.hasCategoryColumn && analysis.hasNumericColumns) {
      suggestions.push({
        type: 'bar',
        reason: 'Categorical comparison',
        confidence: 0.85
      });
    }
    
    // Parts of whole → Pie chart
    if (analysis.sumToWhole && analysis.categoryCount < 10) {
      suggestions.push({
        type: 'pie',
        reason: 'Data represents parts of a whole',
        confidence: 0.8
      });
    }
    
    // Two numeric columns → Scatter
    if (analysis.numericColumnCount >= 2) {
      suggestions.push({
        type: 'scatter',
        reason: 'Explore correlation between variables',
        confidence: 0.7
      });
    }
    
    return suggestions.sort((a, b) => b.confidence - a.confidence);
  }
}
```
