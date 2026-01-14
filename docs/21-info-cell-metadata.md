# INFO() and CELL(): workbook/worksheet metadata requirements

`INFO()` and `CELL()` are Excel “worksheet information” functions. They are **volatile** and depend on **workbook/UI/environment metadata** that is *outside* the pure formula graph (filesystem path, OS info, effective formatting, column widths, etc).

The Rust `formula-engine` intentionally does **not** reach out to the host OS / filesystem / UI. To match Excel, **hosts must inject metadata** into the engine. This document describes:

- which keys are supported today vs planned
- which keys are engine-internal vs host-provided
- what workbook/worksheet metadata is required for Excel-compatible results
- which APIs hosts should call (or which APIs are planned) to supply that metadata

## Status (implementation reality)

As of today:

- `INFO("recalc")` and `INFO("numfile")` are implemented (engine-internal).
- `INFO("system")` is implemented but currently hard-coded to `"pcdos"`.
- Other `INFO()` keys listed below currently return `#N/A` (recognized but not available).
- `CELL("filename")` currently always returns `""` (empty string), matching Excel’s “unsaved workbook” behavior.
- `CELL("protect")`, `CELL("prefix")`, and `CELL("width")` currently return `#N/A` (recognized but not implemented).

Tracking implementation: `crates/formula-engine/src/functions/information/worksheet.rs`.

---

## INFO(type_text)

### Key semantics and data sources

Keys are **trimmed** and **case-insensitive**. Unknown keys return `#VALUE!`.

| Key | Return type | Source | Engine status | Missing metadata behavior |
|---|---|---|---|---|
| `recalc` | text | engine (`CalcSettings.calculation_mode`) | implemented | n/a |
| `numfile` | number | engine (sheet count) | implemented | n/a |
| `system` | text | host | **partially implemented** (currently always `"pcdos"`) | if host does not supply: defaults to `"pcdos"` |
| `directory` | text | host | planned | return `#N/A` if missing |
| `origin` | text | host | planned | return `#N/A` if missing |
| `osversion` | text | host | planned | return `#N/A` if missing |
| `release` | text | host | planned | return `#N/A` if missing |
| `version` | text | host | planned | return `#N/A` if missing |
| `memavail` | number | host | planned | return `#N/A` if missing |
| `totmem` | number | host | planned | return `#N/A` if missing |

#### Notes

- `INFO("recalc")` mirrors Excel’s text:
  - `"Automatic"`
  - `"Automatic except for tables"`
  - `"Manual"`

  In Rust, this is controlled via `formula_engine::Engine::set_calc_settings(...)`.

- `INFO("numfile")` uses the engine’s notion of *sheet count*. Hosts must ensure the engine knows about **all sheets**, not just sheets containing cells. (For example, the WASM `fromJson` path only creates sheets that exist in the JSON payload.)

- `INFO("origin")` is UI-dependent in Excel (top-left visible cell in the active window). It cannot be derived from workbook data alone.

### Host-provided INFO metadata (what to supply)

When implementing the planned keys, hosts should supply a best-effort snapshot of:

- `system`: `"pcdos"` (Windows-style) or `"mac"` (macOS-style). For web, choose one (we default to `"pcdos"` today).
- `directory`: workbook directory / “current directory” string (platform-specific path conventions).
- `origin`: absolute A1 address of the top-left visible cell, including `$` markers (e.g. `"$A$1"`).
- `osversion`: OS version string (exact string should match Excel as closely as feasible on that platform).
- `release` / `version`: application/host version identifiers (to be defined to match Excel behavior once implemented).
- `memavail` / `totmem`: numeric memory values (units to be validated against Excel; treat as bytes unless/until we lock a different convention).

---

## CELL(info_type, [reference])

### Supported keys

Keys are **trimmed** and **case-insensitive**. Unknown keys return `#VALUE!`.

| Key | Return type | Data required | Engine status |
|---|---|---|---|
| `address` | text | sheet name registry | implemented |
| `row` | number | none | implemented |
| `col` | number | none | implemented |
| `contents` | value/text | cell formula/value | implemented |
| `type` | text | cell formula/value | implemented |
| `filename` | text | workbook file metadata + sheet name | **partially implemented** (currently always `""`) |
| `protect` | number | **effective style** (`protection.locked`) | planned |
| `prefix` | text | **effective style** (`alignment.horizontal`) | planned |
| `width` | number | column width + column hidden state | planned |

Other Excel-valid `CELL()` keys (`color`, `format`, `parentheses`, …) are currently recognized but return `#N/A` in this engine.

### Metadata-backed keys (Excel-compatible behavior)

#### `CELL("filename")`

Excel format:

```
path[workbook]sheet
```

Example (Windows-style):

```
C:\Users\me\Documents\[Book1.xlsx]Sheet1
```

Behavior:

- If the workbook is **unsaved** (no filename/path), Excel returns `""` (empty string). The engine matches this today.
- Once saved, the engine needs:
  - workbook directory (may be empty in web contexts; should use the platform’s path separator and match Excel’s display form, typically including a trailing separator)
  - workbook filename (including extension)
  - the referenced sheet name

Where this metadata comes from in this repo today:

- Cross-platform workbook backends return `WorkbookInfo.path` (`@formula/workbook-backend`). Desktop implementations typically set it to an absolute path; the WASM backend currently returns `null` (no filesystem path in the browser).
- For web, hosts generally only know a filename (e.g. from a file picker). When we implement workbook file metadata injection, web hosts should pass `fileName` only and leave `directory` empty.

#### `CELL("protect")` (planned)

Excel returns:

- `1` if the referenced cell is **locked**
- `0` if the referenced cell is **unlocked**

This must be computed from the cell’s **effective style**:

- Use the merged style’s `protection.locked` field.
- Default behavior should match Excel: if protection is unspecified, treat the cell as **locked**.

#### `CELL("prefix")` (planned)

Excel returns a **single-character prefix** describing the effective horizontal alignment.

Planned mapping (must match Excel exactly when implemented):

| Effective alignment | `CELL("prefix")` |
|---|---|
| left | `'` |
| right | `"` |
| center | `^` |
| fill | `\` |
| general/other/unspecified | `""` |

This must use the cell’s **effective alignment** (layered style merge), not just a per-cell format.

#### `CELL("width")` (planned)

`CELL("width")` must reflect:

- the effective **column width** for the referenced cell’s column
- whether the column is **hidden**

Excel uses a somewhat idiosyncratic numeric encoding for column width/hidden; we will document the exact encoding here once implemented. Until then, this key returns `#N/A` in `formula-engine`.

---

## Formatting / style data model (required for `CELL("protect")` and `CELL("prefix")`)

### 1) Style table + `style_id`

The shared model is:

- a **deduplicated style table** (`StyleTable`)
- style references by integer **`style_id`**
- `style_id = 0` is always the default/empty style

Relevant implementations:

- JS (DocumentController): `apps/desktop/src/formatting/styleTable.js` (`StyleTable`, `applyStylePatch`)
- Rust (formula-model): `crates/formula-model/src/style/mod.rs` (`StyleTable`, `Style`)

### 2) Layered style precedence (effective formatting)

Excel-style effective formatting is *layered* (higher precedence overrides lower precedence).

The document model’s precedence order (used by the desktop UI today) is:

```
sheet < col < row < range-run < cell
```

Concrete sources in the JS document model:

- **sheet default**: `sheet.defaultStyleId`
- **column defaults**: `sheet.colStyleIds[col]`
- **row defaults**: `sheet.rowStyleIds[row]`
- **range-run formatting**: `sheet.formatRunsByCol[col]` (see below)
- **cell style**: per-cell `cell.styleId`

Reference implementation: `apps/desktop/src/document/documentController.js::getCellFormat`.

#### Extracting the style layers in JS (DocumentController)

If your host integration already has a `DocumentController` instance, the data needed by the planned `CELL()` keys is available without materializing full effective styles:

- sheet default: `doc.getSheetDefaultStyleId(sheetId)`
- row default: `doc.getRowStyleId(sheetId, row)`
- col default: `doc.getColStyleId(sheetId, col)`
- cell style: `(doc.getCell(sheetId, { row, col }) as any)?.styleId ?? 0`
- range-run style id for a single cell: `doc.getCellFormatStyleIds(sheetId, { row, col })[4]`
- range-run segments for a column: `doc.model.sheets.get(sheetId)?.formatRunsByCol?.get?.(col)` (implementation detail)

### 3) Styles are patches

Each style referenced by an id is treated as a **patch** that can be merged with other patches.

Merge semantics (as implemented in the UI):

- deep merge of objects (e.g. `{ font: { bold:true } }` + `{ font: { italic:true } }`)
- for conflicts, later layers overwrite earlier layers
- `undefined` keys are ignored (no “delete” semantics)

This is implemented in JS by `applyStylePatch(base, patch)` (deep merge).

### 4) Range-run formatting (`formatRunsByCol`)

For large-formatting operations, the model stores compressed “runs” per column:

- `formatRunsByCol: Map<col, FormatRun[]>`
- each `FormatRun` covers rows `[startRow, endRowExclusive)`
- runs are sorted, non-overlapping, and stored only for `styleId !== 0`

See:

- type docs in `apps/desktop/src/document/documentController.js` (`FormatRun`)
- collaboration docs: `docs/06-collaboration.md` (search `formatRunsByCol`)

---

## Column widths and hidden columns (required for `CELL("width")`)

### Column width units

There are two common representations in the codebase:

- **Excel file/model units (“character” width)**: `formula-model` stores `ColProperties.width: Option<f32>` in Excel column-width units (OOXML `<col width="…">` semantics). See `crates/formula-model/src/worksheet.rs`.
- **UI units (pixels at zoom=1)**: the JS document model uses `sheet.view.colWidths` as “base units, zoom=1” and leaves interpretation to the UI grid. See `apps/desktop/src/document/documentController.js` (`SheetViewState.colWidths`).

For Excel-compatible `CELL("width")`, the engine must ultimately work in **Excel column-width units**. Hosts that only have pixel widths must convert (conversion depends on font/DPI; do not assume a fixed ratio unless you intentionally accept approximation).

In the desktop JS document model, column width overrides are accessible via:

- `doc.getSheetView(sheetId).colWidths` (sparse map `{ [colIndex: string]: widthInUiUnits }`)

### Hidden state

Hidden state should be provided explicitly (Excel tracks hidden columns separately from width).

In the Rust model, this is `ColProperties.hidden: bool`. In JS/UI, hidden columns may be represented via outline metadata or view state; when plumbing into the formula engine, prefer an explicit boolean.

---

## Host integration APIs (what to call)

This section documents the intended “wiring points” for hosts. Some calls exist today; others are planned.

### Web/WASM worker (`packages/engine` + `crates/formula-wasm`)

**Exists today**

- Sheet count (`INFO("numfile")`): create all sheets up-front when loading a workbook
  - `WasmWorkbook.fromJson({ sheets: { Sheet1: …, Sheet2: … } })`
  - `WasmWorkbook.fromXlsxBytes(bytes)` (creates all sheets from the XLSX model)
- Calculation mode (`INFO("recalc")`): currently always “Manual” in WASM because `formula_engine::Engine` defaults to manual and the setting is not exposed through the WASM API.

> In practice most web callers use `EngineClient` (`packages/engine/src/client.ts`) rather than calling `WasmWorkbook` directly; the same sheet-count rule applies to the JSON schema passed to `EngineClient.loadWorkbookFromJson(...)`.

**Planned**

Worker/RPC methods to add to `packages/engine` (and corresponding WASM exports on `WasmWorkbook`):

- `setWorkbookFileMetadata({ directory?: string, fileName?: string })`
- `setInfoEnvMetadata({ system?: string, osversion?: string, release?: string, version?: string, memavail?: number, totmem?: number, directory?: string, origin?: string })`
- formatting / styles:
  - `setStyleTable(styles: StylePatch[])` (or a minimal subset needed for `CELL`)
  - `setSheetDefaultStyleId(sheetId, styleId)`
  - `setRowStyleIds(sheetId, updates: Array<{ row: number, styleId: number }>)`
  - `setColStyleIds(sheetId, updates: Array<{ col: number, styleId: number }>)`
  - `setFormatRunsByCol(sheetId, col, runs: Array<{ startRow: number, endRowExclusive: number, styleId: number }>)`
  - `setCellStyleIds(sheetId, updates: Array<{ address: string, styleId: number }>)`
- column widths/hidden:
  - `setColumnProperties(sheetId, updates: Array<{ col: number, width?: number, hidden?: boolean }>)`

### Desktop/Tauri

Desktop hosts generally have access to:

- the absolute workbook path on disk (for `CELL("filename")` and `INFO("directory")`)
- OS version + total/available memory (for `INFO("osversion")`, `INFO("memavail")`, `INFO("totmem")`)
- the UI viewport origin cell (for `INFO("origin")`)
- formatting and column metadata (from the workbook model and/or UI doc model)

**Planned wiring points**

- On workbook open/save:
  - update workbook file metadata (directory + filename)
  - trigger a recalculation so dependent `INFO/CELL` formulas update
- On viewport scroll:
  - update `INFO("origin")`
  - trigger a recalculation (or treat `INFO()` as depending on a “view state” version counter)
- On formatting edits / column resize / hide/unhide:
  - update style/column metadata
  - trigger a recalculation so `CELL("protect")` / `CELL("prefix")` / `CELL("width")` update

---

## Implementation TODOs (for contributors)

These are the concrete code areas that will need changes to fully support the planned keys:

1. **Plumb host metadata into evaluation**:
   - extend `formula_engine::functions::FunctionContext` to expose env/workbook/style/column metadata needed by `INFO/CELL`
2. **Implement keys in `formula-engine`**:
   - `crates/formula-engine/src/functions/information/worksheet.rs`
3. **Expose setters in WASM + worker protocol**:
   - add `WasmWorkbook` methods in `crates/formula-wasm/src/lib.rs`
   - add worker RPC methods/types in `packages/engine/src/*`
4. **Decide + test `CELL("width")` encoding**:
   - lock the exact Excel encoding (width + hidden) in compatibility tests before shipping
