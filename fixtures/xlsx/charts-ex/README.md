# `fixtures/xlsx/charts-ex/`

This directory contains **parser-focused ChartEx workbooks**.

Unlike the lightweight ChartEx stubs under `fixtures/charts/xlsx/` (which may
include a `chartEx` part but omit the real `cx:*Chart` plot/series structure),
these fixtures intentionally include:

- `xl/charts/chartEx1.xml` with a real ChartEx shape:
  - `<cx:plotArea>` and a concrete `<cx:*Chart>` element (e.g. `<cx:waterfallChart>`)
  - series data with formulas (`<cx:f>`) and cached points (`<cx:strCache>`,
    `<cx:numCache>`) so the parser can extract `SeriesModel` values without
    evaluating worksheet formulas.
- `xl/charts/chart1.xml` + `xl/charts/_rels/chart1.xml.rels` referencing the
  ChartEx part via the `.../relationships/chartEx` relationship type.
- `xl/charts/_rels/chartEx1.xml.rels` (empty but present) so OPC graphs match
  what Excel typically emits.

These fixtures are exercised by `crates/formula-xlsx/tests/chart_ex_detection.rs`.

