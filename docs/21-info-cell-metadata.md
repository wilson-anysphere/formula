# INFO() and CELL(): workbook/worksheet metadata requirements

`INFO()` and `CELL()` are Excel “worksheet information” functions. They are **volatile** and depend on **workbook/UI/environment metadata** that is *outside* the pure formula graph (filesystem path, OS info, effective formatting, column widths, etc).

The Rust `formula-engine` intentionally does **not** reach out to the host OS / filesystem / UI. To match Excel, **hosts must inject metadata** into the engine. This document describes:

- which keys are supported today (and what is still missing vs Excel)
- which keys are engine-internal vs host-provided
- what workbook/worksheet metadata is required for Excel-compatible results
- which APIs hosts should call to supply that metadata

## Status (implementation reality)

As of today, the following keys are implemented in `formula-engine`:

- `INFO()` supports the following keys:
  - engine-internal: `INFO("recalc")`, `INFO("numfile")`
  - host-provided (via `EngineInfo`): `INFO("system")`, `INFO("directory")`, `INFO("osversion")`, `INFO("release")`, `INFO("version")`, `INFO("memavail")`, `INFO("totmem")`
    - `system` defaults to `"pcdos"` when unset (Excel-like).
    - `directory` prefers `EngineInfo.directory`; otherwise it falls back to workbook file metadata (`setWorkbookFileMetadata`) and returns `#N/A` unless both a non-empty directory and filename are known.
    - `osversion` / `release` / `version` / `memavail` / `totmem` return `#N/A` when unset.
  - per-sheet view metadata (via `setSheetOrigin` / `Engine::set_sheet_origin`): `INFO("origin")` (defaults to `"$A$1"` when unset)
- `CELL("filename")` returns `""` (empty string) until the host supplies workbook file metadata, matching Excel’s “unsaved workbook” behavior.
- `CELL("protect")` and `CELL("prefix")` are implemented based on the cell’s **effective style** (layered style resolution), matching Excel semantics:
  - `protect`: `1` if the effective style is locked (default), `0` if unlocked.
  - `prefix`: single-character alignment prefix (`'`, `"`, `^`, `\`) or `""` for general/unspecified.
- `CELL("format")`, `CELL("color")`, and `CELL("parentheses")` are implemented based on the cell’s
  **effective number format string** (using `formula-format`) and **ignore conditional formatting**.
- `CELL("width")` is implemented with Excel-compatible encoding:
  - consults per-column metadata (`ColProperties.width` / `ColProperties.hidden`) and the sheet default width
  - returns an encoded number:
    - integer part: `floor(widthChars)`
    - fractional marker: `+ 0.0` when the column uses the sheet default width, `+ 0.1` when it has an explicit per-column width override
  - returns `0` when the column is hidden

  `widthChars` is the Excel column-width unit (OOXML `col/@width`). When unset, the engine falls back to Excel’s standard `8.43` width, which encodes as `8.0`.

Tracking implementation: `crates/formula-engine/src/functions/information/worksheet.rs`.

---

## INFO(type_text)

### Key semantics and data sources

Keys are **trimmed** and **case-insensitive**. Unknown keys return `#VALUE!`.

| Key | Return type | Source | Engine status | Missing metadata behavior |
|---|---|---|---|---|
| `recalc` | text | engine (`CalcSettings.calculation_mode`) | implemented | n/a |
| `numfile` | number | engine (sheet count) | implemented | n/a |
| `system` | text | host (`EngineInfo.system`) | implemented | defaults to `"pcdos"` when unset |
| `directory` | text | host (`EngineInfo.directory`) **or** workbook file metadata | implemented | returns `#N/A` unless `EngineInfo.directory` is set, or workbook file metadata includes both a non-empty directory and filename |
| `origin` | text | host (per-sheet origin cell) | implemented | defaults to `"$A$1"` when unset |
| `osversion` | text | host (`EngineInfo.osversion`) | implemented | return `#N/A` when unset |
| `release` | text | host (`EngineInfo.release`) | implemented | return `#N/A` when unset |
| `version` | text | host (`EngineInfo.version`) | implemented | return `#N/A` when unset |
| `memavail` | number | host (`EngineInfo.memavail`) | implemented | return `#N/A` when unset |
| `totmem` | number | host (`EngineInfo.totmem`) | implemented | return `#N/A` when unset |

#### Notes

- `INFO("recalc")` mirrors Excel’s text:
  - `"Automatic"`
  - `"Automatic except for tables"`
  - `"Manual"`

  In Rust, this is controlled via `formula_engine::Engine::set_calc_settings(...)`.

- `INFO("numfile")` uses the engine’s notion of *sheet count*. Hosts must ensure the engine knows about **all sheets**, not just sheets containing cells. (For example, the WASM `fromJson` path only creates sheets that exist in the JSON payload.)

- `INFO("origin")` is UI-dependent in Excel (top-left visible cell in the active window). It cannot be derived from workbook data alone. The engine models this as a **per-sheet origin cell**.

### Host-provided INFO metadata (what to supply)

Hosts should supply a best-effort snapshot of:

- `system`: `"pcdos"` (Windows-style) or `"mac"` (macOS-style). When unset, the engine defaults to `"pcdos"`.
- `directory`:
  - host override via `EngineInfo.directory`
  - otherwise supplied implicitly via workbook file metadata (`setWorkbookFileMetadata`) when the host has an OS path (desktop); if only a filename is known (common on web), `INFO("directory")` remains `#N/A`
  - the engine returns a trailing path separator (e.g. `"/tmp/"`, `"C:\\Dir\\"`) to match Excel
- `origin`: set per-sheet via `setSheetOrigin(sheet, originA1)`; the engine always returns an absolute A1 string with `$` markers (e.g. `"$C$5"`). When unset, the engine defaults to `"$A$1"`.
- `osversion`: OS version string (exact string should match Excel as closely as feasible on that platform).
- `release` / `version`: application/host version identifiers.
- `memavail` / `totmem`: numeric memory values. In WASM, these must be finite numbers (`NaN`/`Infinity` are rejected). Units are host-defined (bytes recommended).

---

## CELL(info_type, [reference])

### Supported keys

Keys are **trimmed** and **case-insensitive**. Unknown keys return `#VALUE!`.

For local references, if the `reference` argument resolves outside the sheet’s configured dimensions (see `setSheetDimensions` / `Engine::set_sheet_dimensions`), the engine returns `#REF!` (Excel-like).

| Key | Return type | Data required | Engine status |
|---|---|---|---|
| `address` | text | sheet name registry | implemented |
| `row` | number | none | implemented |
| `col` | number | none | implemented |
| `contents` | value/text | cell formula/value | implemented |
| `type` | text | cell formula/value | implemented |
| `format` | text | **effective number format string** | implemented (uses `formula-format`, ignores conditional formatting) |
| `color` | number | **effective number format string** | implemented (uses `formula-format`, ignores conditional formatting) |
| `parentheses` | number | **effective number format string** | implemented (uses `formula-format`, ignores conditional formatting) |
| `filename` | text | workbook file metadata + sheet name | implemented (returns `""` until metadata is set) |
| `protect` | number | **effective style** (`protection.locked`) | implemented |
| `prefix` | text | **effective style** (`alignment.horizontal`) | implemented |
| `width` | number | column width + column hidden state | implemented (encodes `floor(widthChars) + 0.0/0.1`; returns `0` when hidden) |
Other Excel-valid `CELL()` keys that are not listed above are currently **unsupported** and return `#VALUE!` (unknown `info_type`).

### Metadata-backed keys (Excel-compatible behavior)

#### `CELL("address")`

`CELL("address")` returns an **absolute** A1 reference like `"$C$5"`.

When the returned address needs a sheet prefix, the engine emits:

```text
sheet!$A$1
```

Behavior:

- If the reference is on the **current sheet**, the result is just `"$A$1"` (no `sheet!` prefix).
- If the reference is on a **different sheet**, the engine prefixes the address with that sheet’s **display name**
  (tab name), as configured via `setSheetDisplayName` (or `renameSheet`).
- For **external references**, the engine prefixes with the canonical external sheet key (e.g. `"[Book.xlsx]Sheet1"`).
  Because that key contains `[]`, it is quoted in the output:
  - `'[Book.xlsx]Sheet1'!$A$1`
- The sheet prefix is quoted using Excel formula rules when required (spaces, punctuation, reserved names like
  `TRUE`/`FALSE`, names that look like cell refs, etc). Apostrophes are escaped by doubling:
  - `'Other Sheet'!$A$1`
  - `'O''Brien'!$A$1`

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
- The engine treats the workbook as “saved” only once a **non-empty** `filename` is known; supplying a directory alone is not enough.
- When only a filename is known (common on web), `CELL("filename")` returns:
  - `[Book.xlsx]Sheet1` (no directory prefix)
- In this case, `INFO("directory")` returns `#N/A` (unless overridden via `EngineInfo.directory`).
- When a directory is also known, the engine returns Excel’s `path[workbook]sheet` shape, ensuring the directory has a trailing separator:
  - `C:\Dir\[Book.xlsx]Sheet1`
  - `/dir/[Book.xlsx]Sheet1`
- `CELL("filename", reference)` uses the referenced cell’s sheet name component.
  - For local sheets, this is the sheet’s **display name** (tab name) as configured via `setSheetDisplayName` (or `renameSheet`).
  - For external references, the engine returns the canonical external sheet key (e.g. `"[Book.xlsx]Sheet1"`), which already includes the workbook + sheet.
- The sheet name portion is **not quoted** in the output (even if the sheet name contains spaces), matching Excel:
  - `[Book.xlsx]Other Sheet` (not `[Book.xlsx]'Other Sheet'`)

In this repo, hosts can inject this metadata via:

- WASM/worker (`@formula/engine`): `EngineClient.setWorkbookFileMetadata(directory, filename)`

Where this metadata comes from in this repo today:

- Cross-platform workbook backends return `WorkbookInfo.path` (`@formula/workbook-backend`). Desktop implementations typically set it to an absolute path; the WASM backend currently returns `null` (no filesystem path in the browser).
- For web, hosts generally only know a filename (e.g. from a file picker). Web hosts should pass `filename` only and leave `directory` empty/`null`.

#### `CELL("protect")`

Excel returns:

- `1` if the referenced cell is **locked**
- `0` if the referenced cell is **unlocked**

This is computed from the cell’s **effective style**:

- Uses the merged style’s `protection.locked` field.
- Default behavior matches Excel: if protection is unspecified, treat the cell as **locked**.
- The result reflects formatting only and does **not** depend on whether sheet protection is enabled.
- If the reference points outside the sheet’s configured dimensions (see `Engine::set_sheet_dimensions`),
  the engine returns `#REF!` (mirrors `get_cell_value` bounds behavior).

#### `CELL("prefix")`

Excel returns a **single-character prefix** describing the effective horizontal alignment.

Mapping (Excel-compatible):

| Effective alignment | `CELL("prefix")` |
|---|---|
| left | `'` |
| right | `"` |
| center | `^` |
| fill | `\` |
| general/other/unspecified | `""` |

This uses the cell’s **effective alignment** (layered style merge), not just a per-cell format.
- If the reference points outside the sheet’s configured dimensions, the engine returns `#REF!`.

#### `CELL("width")`

`CELL("width")` must reflect:

- the effective **column width** for the referenced cell’s column
- whether the column is **hidden**

Current behavior (Excel encoding):

- If the referenced column is hidden (`ColProperties.hidden = true`): returns `0`.
- Otherwise returns:
  - `floor(widthChars) + 0.0` when the column uses the sheet default width
  - `floor(widthChars) + 0.1` when the column has an explicit per-column width override

`widthChars` is stored in Excel column-width units (OOXML `<col width="…">` semantics). When unset, the engine falls back to the Excel standard `8.43`.
- If the reference points outside the sheet’s configured dimensions, the engine returns `#REF!`.

#### `CELL("format")` / `CELL("color")` / `CELL("parentheses")`

These keys are computed from the cell’s **effective number format string**, not from the cell’s value.
- If the reference points outside the sheet’s configured dimensions (see `Engine::set_sheet_dimensions`),
  the engine returns `#REF!` (mirrors `get_cell_value` bounds behavior).

- `CELL("format")` returns an Excel format code string (e.g. `"G"`, `"F2"`, `"N0"`, `"C2"`).
- `CELL("color")` returns `1` if the **negative section** of the number format specifies a color (e.g. a second section with a color tag), otherwise `0`.
- `CELL("parentheses")` returns `1` if the **negative section** of the number format uses parentheses for negatives, otherwise `0`.

Example number format with a colored negative section:

```text
0;[Red] (0)
```

Notes:

- The engine first checks for an **explicit per-cell number format override** (if one exists). In Rust, hosts can set this via `Engine::set_cell_number_format(sheet, addr, pattern)`. This override takes precedence over style layers.
- Spilled output cells (dynamic arrays) inherit their number format from the spill origin cell.
- Otherwise, the number format is resolved from the effective formatting layers (Excel-like precedence): `cell (non-zero) > range-run > row > col > sheet default > 0 (General)`.
  - Styles that do not specify `number_format` are treated as “inherit”, so lower-precedence layers can contribute the number format.
- Conditional formatting is ignored.
- When number format metadata is unavailable (no style table, or external workbook refs), the engine falls back to General semantics:
  - `CELL("format")` → `"G"`
  - `CELL("color")` → `0`
  - `CELL("parentheses")` → `0`

---

## Formatting / style data model (required for `CELL("protect")`, `CELL("prefix")`, `CELL("format")`, `CELL("color")`, `CELL("parentheses")`)

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

If your host integration already has a `DocumentController` instance, the data needed by formatting-backed `CELL()` keys is available without materializing full effective styles:

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
- `undefined` keys are ignored (patch does not affect the base style)
- `null` values are meaningful and are commonly used by the UI layer to **clear** a field back to its default
  (e.g. `numberFormat: null` to clear to General; `alignment: { horizontal: null }` to clear to General;
  `protection: { locked: null }` to clear back to Excel's default locked=true; `protection: { hidden: null }` to clear back to hidden=false;
  `applyStylePatch(base, null)` resets to `{}`)

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

This repo includes a best-effort conversion helper in `@formula/engine`:

- `pixelsToExcelColWidthChars(pixels)` / `excelColWidthCharsToPixels(widthChars)` (`packages/engine/src/columnWidth.ts`)

These implement Excel’s default-font conversion (Calibri 11, max digit width 7px, padding 5px) and are an approximation if the workbook’s default font differs.

In the desktop JS document model, column width overrides are accessible via:

- `doc.getSheetView(sheetId).colWidths` (sparse map `{ [colIndex: string]: widthInUiUnits }`)

### Hidden state

Hidden state should be provided explicitly (Excel tracks hidden columns separately from width).

In the Rust model, this is `ColProperties.hidden: bool`. In JS/UI, hidden columns may be represented via outline metadata or view state; when plumbing into the formula engine, prefer an explicit boolean.

---

## Host integration APIs (what to call)

This section documents the “wiring points” for hosts.

### Web/WASM worker (`packages/engine` + `crates/formula-wasm`)

**Exists today (public API: `EngineClient` from `packages/engine/src/client.ts`)**

- Sheet count (`INFO("numfile")`): create all sheets up-front when loading a workbook
  - `WasmWorkbook.fromJson({ sheets: { Sheet1: …, Sheet2: … } })`
  - `WasmWorkbook.fromXlsxBytes(bytes)` (creates all sheets from the XLSX/XLSM model)
 
- Sheet display names (affects `CELL("address")` and the sheet component of `CELL("filename")`):
  - `EngineClient.setSheetDisplayName(sheetId, name)`
  - This is metadata-only: it does **not** rewrite stored formulas and does **not** change the engine-facing sheet id/key.
  - If your engine-facing sheet ids/keys are not valid Excel tab names (e.g. UUIDs), the engine will still generate an Excel-valid default display name (typically `Sheet{n}`). Use `setSheetDisplayName` to control what `CELL("address")`/`CELL("filename")` emit.
  - For Excel-like rename semantics (rewrite formulas that reference a sheet), use `EngineClient.renameSheet(oldName, newName)` when available.

- Sheet dimensions (worksheet bounds; affects `CELL(..., reference)` out-of-bounds `#REF!` behavior and whole-row/whole-column references like `A:A` / `1:1`):
  - `EngineClient.setSheetDimensions(sheet, rows, cols)`
  - `EngineClient.getSheetDimensions(sheet)` (debugging/introspection)
  - When unset, the engine uses Excel’s default grid size (`1,048,576` rows × `16,384` cols).

- Calculation mode (`INFO("recalc")`): exposed via `EngineClient.getCalcSettings()` / `EngineClient.setCalcSettings()`.
  - Note: the worker protocol typically runs edits in manual mode so JS callers can explicitly request `recalculate()` and receive deterministic value-change deltas.
  - In practice, most mutating WASM APIs preserve **explicit-recalc semantics even when the workbook calcMode is Automatic** (they are wrapped in `with_manual_calc_mode` in `crates/formula-wasm`), so callers should assume they need to call `recalculate()` to observe updated `INFO()`/`CELL()` results.

- INFO environment metadata (`INFO("system")`, `INFO("directory")`, `INFO("osversion")`, `INFO("release")`, `INFO("version")`, `INFO("memavail")`, `INFO("totmem")`):
  - `EngineClient.setEngineInfo({ system, directory, osversion, release, version, memavail, totmem })`
    - The worker/WASM implementation treats this as a **patch**: fields not present in the object are left unchanged.
    - `null` / empty string clears string values
    - `memavail` / `totmem` must be finite numbers
- Per-sheet origin cell (`INFO("origin")`, preferred):
  - `EngineClient.setSheetOrigin(sheet, originA1 | null)`
  - `originA1` should be an in-bounds A1 address (`"C5"` or `"$C$5"`); the engine returns absolute A1 with `$`.
  - Passing `null` (or `""` via the worker RPC layer) clears the origin and restores the default `"$A$1"`.
  - Compatibility note:
    - `EngineClient.setInfoOriginForSheet(sheet, originA1|null)` is a legacy alias for `setSheetOrigin` (A1-only, validated).
    - `EngineClient.setInfoOrigin(origin|null)` sets a workbook-level legacy string fallback used when no per-sheet origin is set. If it parses as A1 it is normalized to absolute A1; otherwise it is returned verbatim.

- Workbook file metadata (`CELL("filename")`, `INFO("directory")` fallback):
  - `EngineClient.setWorkbookFileMetadata(directory, filename)`
    - `CELL("filename")` returns `""` until `filename` is known (Excel unsaved behavior)
    - `INFO("directory")` returns `#N/A` unless `EngineInfo.directory` is set, or both `filename` and a non-empty `directory` are known
    - Note: in older/minimal WASM builds, `WasmWorkbook.setWorkbookFileMetadata` may be missing; the worker treats this as a no-op.

- Formatting metadata (`CELL("protect")`, `CELL("prefix")`, `CELL("format")`, `CELL("color")`, `CELL("parentheses")`):
  - `EngineClient.internStyle(stylePatch)` → `styleId`
  - `EngineClient.setSheetDefaultStyleId(sheet, styleId|null)`
  - `EngineClient.setRowStyleId(sheet, row, styleId|null)`
  - `EngineClient.setColStyleId(sheet, col, styleId|null)`
  - range-run formatting layer (DocumentController `formatRunsByCol`):
    - `EngineClient.setFormatRunsByCol(sheet, col, runs)` (preferred)
    - `EngineClient.setColFormatRuns(sheet, col, runs)` (legacy alias used by some sync surfaces)
  - `EngineClient.setCellStyleId(address, styleId, sheet)`

  Notes:
  - These keys use **number format strings** and **layered styles** (sheet/col/row/range-run/cell); conditional formatting is ignored.
  - Style ids are workbook-global; `0` is always the default/empty style.

- Column metadata (`CELL("width")`):
  - `EngineClient.setSheetDefaultColWidth(sheet, widthChars)` to set the sheet default width used for columns without explicit overrides (`null` clears back to Excel’s standard `8.43`).
  - `EngineClient.setColWidthChars(sheet, col, widthChars)` (preferred) or `EngineClient.setColWidth(col, widthChars, sheet)`
    - widths are in Excel “character” units (OOXML `col/@width`), not pixels
  - `EngineClient.setColHidden(col, hidden, sheet)` to set the explicit hidden flag

> In practice most web callers use `EngineClient` (`packages/engine/src/client.ts`) rather than calling `WasmWorkbook` directly; the same sheet-count rule applies to the JSON schema passed to `EngineClient.loadWorkbookFromJson(...)`.

### Desktop/Tauri

Desktop hosts generally have access to:

- the absolute workbook path on disk (for `CELL("filename")` and `INFO("directory")`)
- OS version + total/available memory (for `INFO("osversion")`, `INFO("memavail")`, `INFO("totmem")`)
- the UI viewport origin cell (for `INFO("origin")`)
- formatting and column metadata (from the workbook model and/or UI doc model)

**Suggested wiring points**

- On workbook open/save:
  - update workbook file metadata (directory + filename)
  - trigger a recalculation so dependent `INFO/CELL` formulas update
- On viewport scroll:
  - update `INFO("origin")` by calling `setSheetOrigin(sheet, originA1)`
  - trigger a recalculation (or treat `INFO()` as depending on a “view state” version counter)
- On formatting edits / column resize / hide/unhide:
  - update style/column metadata
  - trigger a recalculation so formatting/width-backed `CELL()` keys update (`protect`, `prefix`, `format`, `color`, `parentheses`, `width`)

---

## Implementation notes (for contributors)

Most of the metadata plumbing described above is implemented. Key code locations:

- Formula implementations: `crates/formula-engine/src/functions/information/worksheet.rs`
- Engine host metadata setters: `crates/formula-engine/src/engine.rs`
- WASM exports: `crates/formula-wasm/src/lib.rs`
- Worker RPC protocol + dispatch: `packages/engine/src/protocol.ts` and `packages/engine/src/engine.worker.ts`

Remaining TODOs / future work:

1. **Conditional formatting**:
   - Excel’s `CELL("color")` semantics involve *format strings*, not conditional formatting, but other potential `CELL()` keys (and future UI parity work) may require modeling conditional formats.
2. **Additional Excel `CELL()` keys**:
   - Implement more `CELL(info_type)` variants (e.g. `row`, `col` are done; others are still missing).
3. **Keep metadata encoding tests stable**:
   - `CELL("width")` integer+flag encoding (default/custom/hidden)
