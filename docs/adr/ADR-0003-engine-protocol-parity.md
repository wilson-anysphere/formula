# ADR-0003: Engine protocol boundary (A1 coordinates, sheet identity, formula normalization)

- **Status:** Accepted
- **Date:** 2026-01-11

## Context

Today we have three overlapping “engine” surfaces that are close enough to look interchangeable, but different enough to drift:

1. **Desktop (Tauri) workbook backend**: row/col + `sheetId` commands (e.g. `get_range`, `set_cell`) and returns `display_value` strings.
   - TypeScript client: `apps/desktop/src/tauri/workbookBackend.ts`
   - Sync glue from `DocumentController` → Tauri: `apps/desktop/src/tauri/workbookSync.ts`
2. **Web/WASM Worker engine**: A1 address + “sheet” RPC (currently treated as a sheet name) returning scalar JSON (`CellScalar`).
   - TS package: `packages/engine/` (`EngineClient`, `EngineWorker`, `protocol.ts`)
3. **Formula text normalization drift**: `formula-model` stores formulas **without** a leading `'='`, while UI/editor workflows generally use the display form **with** `'='`.
   - Canonical helpers: `crates/formula-model/src/formula_text.rs`

Without a single source of truth, we end up duplicating conversions (row/col ↔ A1, `sheetId` ↔ sheet name, with/without `'='`) in multiple places and create subtle incompatibilities (especially around renames and round-trips).

## Decision

### 1) Canonical coordinate protocol at the JS ↔ engine boundary: **A1 strings**

At the boundary where JavaScript calls into “the engine” (Worker/WASM or a host adapter), **cell and range addresses are represented as A1 strings**:

- `address: "B3"`
- `range: "A1:C10"` (inclusive range, Excel semantics)

Row/col coordinates remain the **UI/internal** representation and are converted at the boundary using shared helpers:

- `@formula/spreadsheet-frontend/a1` (`toA1`, `fromA1`, `range0ToA1`)

**Why A1**

- Matches Excel’s user-visible reference style and the syntax embedded inside formulas.
- Avoids 0-based/1-based ambiguity leaking across protocol boundaries.
- Keeps the Worker RPC human-readable for debugging/telemetry.
- Aligns with the existing `@formula/engine` API and shared web preview (`packages/spreadsheet-frontend`).

**Normalization rule**

- Engine RPC accepts A1 addresses/ranges with optional `$` markers (e.g. `$A$1`), but call sites should prefer canonical non-absolute forms (`A1`, `A1:B2`) when addressing cells/ranges out-of-band from formula text.

### 2) Sheet identity: **separate `sheetId` (stable) from `sheetName` (display/formula)**

We explicitly model two concepts:

- `sheetId`: stable, opaque identifier used by the document model / collaboration / UI routing.
- `sheetName`: user-visible worksheet name (Excel semantics) used in:
  - formula text (`'My Sheet'!A1`)
  - file formats (XLSX workbook sheet names)

**Normalization rules**

- `sheetId`:
  - treated as **opaque** and **case-sensitive**
  - must be unique within a workbook
  - does **not** change on rename
- `sheetName`:
  - trimmed for display (`name.trim()`)
  - unique **case-insensitively** within a workbook (Excel behavior)
  - when embedded in formulas, follows Excel quoting rules:
    - quote with `'` when needed
    - escape internal `'` by doubling (`O'Brien` → `'O''Brien'`)

**Engine protocol rule (target state)**

All engine RPC methods that take a sheet identifier take a `sheetId` (even if the field is currently named `sheet` in TS types). Engines must maintain a registry mapping:

```
sheetId -> sheetName
```

This registry is what the formula parser/evaluator uses to resolve `SheetName!A1` references, and what the serializer uses when emitting formula text.

**Important: current implementation constraint**

The current WASM engine (`crates/formula-wasm`, backed by `crates/formula-engine`) is sheet-name keyed and does not yet implement a `sheetId -> sheetName` registry. Until that exists, web and desktop parity assumes **`sheetId === sheetName`** for any sheet that participates in cross-sheet formulas.

This is acceptable for the initial iteration because:

- imported workbooks currently derive ids from names (`apps/desktop/src-tauri/src/file_io.rs::ensure_sheet_ids`), and
- the web preview uses `Sheet1`, `Sheet2`, … as both id and name.

Renaming sheets without rebuilding the engine state is explicitly deferred (see Non-goals).

### 3) Formula string normalization across layers

We standardize formula text as follows:

| Layer | Stored form | Example |
|------|-------------|---------|
| UI/editor + engine inputs | **display form** (leading `'='`) | `=SUM(A1:A3)` |
| `formula-model` (Rust) | **canonical form** (no leading `'='`) | `SUM(A1:A3)` |
| XLSX SpreadsheetML `<f>` | **no leading `'='`** | `<f>SUM(A1:A3)</f>` |

**Where conversions happen**

- **Model boundary (Rust)**: use the canonical helpers:
  - `formula_model::formula_text::normalize_formula_text` (strip `'='`, trim, empty → `None`)
  - `formula_model::formula_text::display_formula_text` (ensure leading `'='` for UI/engine)
- **Desktop file I/O**:
  - On load, formulas from `formula-model` are converted to display form when building the in-memory app workbook (`apps/desktop/src-tauri/src/file_io.rs`).
  - On save, formulas must be converted back to canonical/no-`'='` before writing `<f>` parts (owned by `formula-xlsx` + `formula-model`).
- **JS → engine adapters**:
  - `apps/desktop/src/tauri/workbookSync.ts` and `packages/engine/src/documentControllerSync.ts` currently ensure a leading `'='` before calling an engine/backend.
  - Parity plan is to dedupe this logic so every platform applies *exactly* the same normalization.

### 4) Protocol parity plan (desktop vs web)

Goal: shared UI code should talk to **one** engine-shaped API, independent of platform.

**Chosen strategy: adapter layer, not mass call-site rewrites**

1. **Keep `@formula/engine` (`packages/engine`) as the canonical JS engine API** (A1 + sheet selector + scalar JSON).
2. **Add a desktop adapter that implements the same API on top of Tauri**:
   - Converts A1/range → row/col rectangles for `get_range`/`set_range`.
   - Converts `sheetId` → backend `sheet_id` and uses workbook metadata for mapping when needed.
   - Bridges return types (see note below on display formatting).

Concrete touchpoints:

- Canonical API: `packages/engine/src/client.ts` (`EngineClient`)
- Desktop transport today:
  - `apps/desktop/src/tauri/workbookBackend.ts` (Tauri `invoke` calls)
  - `apps/desktop/src/tauri/workbookSync.ts` (DocumentController delta batching)
- Web transport today:
  - `packages/engine/src/worker/*` + `crates/formula-wasm/`

**Note on display formatting parity**

Desktop currently returns `display_value` strings from the host (`apps/desktop/src-tauri/src/commands.rs`), while the Worker/WASM engine returns typed scalar JSON.

The canonical protocol is **typed scalar values**; formatting into display strings is a separate concern (eventually shared via `formula-format` / WASM). In the first iteration the desktop adapter may temporarily surface formatted values as strings to preserve existing UI behavior, but new shared UI code should avoid baking in host-only `display_value`.

### 5) Non-goals (first iteration)

- **Full XLSX OPC part preservation in the web target**. The WASM engine can load `.xlsx`/`.xlsm` bytes via `crates/formula-wasm::fromXlsxBytes`, but it only imports the workbook model (values/formulas/basic metadata) and does not preserve arbitrary OPC parts on round-trip.
- **VBA execution / macro enablement in web**.
- **Chart/pivot fidelity parity** between web and desktop engines.
- **Sheet rename parity without engine rebuild** (requires a sheet registry + formula rewrite plumbing).

## Migration steps (from today → desired end state)

1. **Introduce/standardize an engine adapter boundary**:
   - Desktop: wrap `apps/desktop/src/tauri/workbookBackend.ts` behind an `EngineClient`-shaped adapter that speaks A1.
   - Web: continue using `createEngineClient()` from `packages/engine`.
2. **Unify DocumentController → engine syncing**:
   - Consolidate the duplicated “delta batching + A1 conversion + formula normalization” logic from:
     - `apps/desktop/src/tauri/workbookSync.ts`
     - `packages/engine/src/documentControllerSync.ts`
   - into a single shared helper (location TBD) so both platforms produce identical engine inputs.
3. **Make sheet identity explicit in the protocol**:
   - Rename TS fields from `sheet` → `sheetId` in `packages/engine/src/protocol.ts` (and downstream types) once adapters exist.
   - Add a `sheet registry` sync step so WASM can decouple `sheetId` and `sheetName` without requiring equality.
4. **Centralize formula text conversions**:
   - Replace ad-hoc JS `normalizeFormulaText()` helpers with a shared implementation that matches
     `crates/formula-model/src/formula_text.rs` semantics (including edge cases like bare `"="`).
