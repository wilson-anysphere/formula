# AI Tab Completion (design + extension guide)

Formula’s tab completion is built to feel instant while the user is typing. The desktop UX target is **<100ms end-to-end** on each keystroke, which is why the system is primarily local heuristics plus an optional Cursor backend completion, all wrapped in aggressive caching and hard limits.

This doc describes:
- How `TabCompletionEngine` produces suggestions (rule-based, pattern-based, backend-based)
- How schema and structured references are integrated (`SchemaProvider`)
- How range suggestions work today (`rangeSuggester.js`) and how to extend them safely
- Desktop UI expectations (what can/can’t be rendered as ghost text)
- Where the tests live and how to run them

## Code map

**Core engine (`@formula/ai-completion`):**
- `packages/ai-completion/src/tabCompletionEngine.js` — `TabCompletionEngine`
- `packages/ai-completion/src/lruCache.js` — `LRUCache`
- `packages/ai-completion/src/formulaPartialParser.js` — `parsePartialFormula()` (partial formula parsing)
- `packages/ai-completion/src/rangeSuggester.js` — `suggestRanges()` (contiguous range heuristics)
- `packages/ai-completion/src/patternSuggester.js` — `suggestPatternValues()` (non-formula “value” completion)
- `packages/ai-completion/src/cursorTabCompletionClient.js` — Cursor backend client (optional)

**Desktop integration:**
- `apps/desktop/src/ai/completion/formulaBarTabCompletion.ts` — connects the engine to the formula bar
- `apps/desktop/src/formula-bar/FormulaBarModel.ts` — accepts suggestions + derives ghost text
- `apps/desktop/src/formula-bar/FormulaBarView.ts` — renders ghost text + preview

---

## 1) Completion pipeline (`TabCompletionEngine`)

`TabCompletionEngine.getSuggestions(context, { previewEvaluator? })` is the single entrypoint used by UI surfaces.

### Inputs

The engine consumes a `CompletionContext` (see `packages/ai-completion/src/tabCompletionEngine.js`):
- `currentInput`: full draft string (formula or value)
- `cursorPosition`: caret position (0…`currentInput.length`)
- `cellRef`: `{row, col}` or an A1 string
- `surroundingCells`: an object with:
  - `getCellValue(row, col, sheetName?)` → any (should be cheap + synchronous)
  - optional `getCacheKey()` → string (see caching section)

### Stages

1. **Normalize** input + cursor
   - Cursor is clamped into the valid range (`clampCursor()`).

2. **Cache lookup**
   - A cache key is computed via `buildCacheKey()` (details below).
   - If present, cached **base suggestions** are reused.

3. **Compute base suggestions (in parallel)**
   - The engine parses the draft with `parsePartialFormula()`, then runs three sources concurrently:

   **A. Rule-based (deterministic)**
   - Implemented by `getRuleBasedSuggestions()`.
   - Only active for formulas (`parsed.isFormula`).
   - Covers:
     - **Function name completion** (e.g. `=VLO` → `=VLOOKUP(`) via `suggestFunctionNames()`.
     - **Workbook identifiers** when not in a call (named ranges + structured refs) via `suggestWorkbookIdentifiers()`.
     - **Range completions** when a function argument expects a range (e.g. `=SUM(A` → `=SUM(A1:A10)`) via `suggestRangeCompletions()`.
     - **Argument value hints** for simple arg types (`TRUE/FALSE`, `0/1`, “cell to the left”) via `suggestArgumentValues()`.

   **B. Pattern-based (local value repetition)**
   - Implemented by `getPatternSuggestions()`.
   - Only active for non-formula input (`!parsed.isFormula`).
   - Uses `suggestPatternValues()` to scan nearby cells for repeated strings that match the typed prefix.

    **C. Cursor backend completion (optional)**
    - Implemented by `getCursorBackendSuggestions()`.
    - Only runs when:
      - a `completionClient` is configured, and
      - we are editing a formula (`parsed.isFormula`), and
      - the user is not currently typing a function name prefix (`!parsed.functionNamePrefix`).
    - Calls:
      - `completionClient.completeTabCompletion({ input, cursorPosition, cellA1 })`
    - The request is wrapped in a **hard timeout** (`completionTimeoutMs`, default 100ms; clamped to ≤200ms) so the UI stays responsive even if the backend is slow/unavailable.
    - The built-in Cursor backend client (`CursorTabCompletionClient`) supports:
      - `getAuthHeaders?: () => Record<string,string> | Promise<Record<string,string>>` for environments where cookie-based auth isn't available.
      - `signal?: AbortSignal` on `completeTabCompletion()` for caller-driven cancellation (e.g. abort in-flight requests when the user keeps typing).

4. **Rank + dedupe**
   - Suggestions are merged, then `rankAndDedupe()`:
     - keeps the highest-confidence suggestion per `text`
     - sorts by confidence (then a small type priority)
     - returns the top `maxSuggestions` (default 5)

5. **Optional preview attachment**
   - If `previewEvaluator` is provided, it is called for each `"formula"`/`"range"` suggestion and the result is attached as `suggestion.preview`.
   - The cache stores **base suggestions** (without previews). Previews are best-effort and surface-specific.

### Suggestion shape

Suggestions are plain objects:
- `text`: the **full suggested draft string** (what would be committed if accepted)
- `displayText`: a UI-friendly display snippet (often just the inserted token)
- `type`: `"formula" | "value" | "function_arg" | "range"`
- `confidence`: `0…1`
- optional `preview`: surface-specific preview payload

---

## 1.1 Caching and cache keys (`LRUCache`)

Tab completion is called on most keystrokes. To stay under 100ms, the engine caches base suggestions.

### Cache implementation

`packages/ai-completion/src/lruCache.js` implements a small `Map`-backed LRU:
- default capacity: **200 entries**
- eviction: oldest-first whenever size exceeds max

### Cache key contents

`TabCompletionEngine.buildCacheKey()` returns a JSON string containing:
- `input`: the full input string
- `cursor`: the cursor position
- `cell`: the normalized active cell (`{row,col}`)
- `surroundingKey`: `surroundingCells.getCacheKey()` (or `""`)
- `schemaKey`: `schemaProvider.getCacheKey()` (or `""`)

These five fields are the core “correctness contract” for cached suggestions: if any of them changes, suggestions may legitimately change.

### Extension guidance (cache keys)

If you’re integrating the engine into a new surface:
- Implement `surroundingCells.getCacheKey()` using a **cheap version counter**, not by hashing large neighborhoods.
  - Desktop uses `${sheetId}:${cellsVersion}` (see `apps/desktop/src/ai/completion/formulaBarTabCompletion.ts`).
- If you provide a `SchemaProvider`, implement `schemaProvider.getCacheKey()` so suggestions refresh when the schema changes (new sheets, renamed tables, etc).
  - Keep the key small and stable (e.g. names + a schema version).

---

## 2) Schema integration (`SchemaProvider`)

The engine can optionally use workbook schema to suggest:
- named ranges (`Revenue`, `salesData`, …)
- sheet names for sheet-qualified ranges (`Sheet2!A…`)
- structured references for tables (`Table1[Amount]`)

### `SchemaProvider` interface

Defined (as JSDoc types) in `packages/ai-completion/src/tabCompletionEngine.js`:
- `getNamedRanges?: () => Array<{ name: string, range?: string }>`
- `getSheetNames?: () => string[]`
- `getTables?: () => Array<{ name: string, columns: string[], sheetName?, startRow?, startCol?, endRow?, endCol? }>`
- `getCacheKey?: () => string`

All methods may return values synchronously or as a `Promise`. The engine uses a defensive wrapper (`safeProviderCall`) and treats errors as empty schema.

### Named ranges

Where they’re suggested:
- inside functions expecting ranges (`suggestSchemaRanges()`)
- when typing an identifier outside a function call (`suggestWorkbookIdentifiers()`)

Key behavior:
- Matching is prefix-based and case-insensitive.
- Completions intentionally preserve the user-typed prefix so the suggestion can be rendered/applied as a “pure insertion” at the caret (see UI section).

### Structured references (tables)

`getTables()` powers structured references via `suggestStructuredRefs()`:
- Primary suggestion: `TableName[ColumnName]`
- Lower-confidence alternative: `TableName[[#All],[ColumnName]]`

Notes:
- The suggestion engine only requires `{ name, columns }`.
- Desktop preview evaluation becomes much better if table coordinates are provided (`sheetName`, `startRow/startCol/endRow/endCol`) because previews can rewrite structured references into A1 ranges (see `rewriteStructuredReferences()` in `apps/desktop/src/ai/completion/formulaBarTabCompletion.ts`).

### Sheet-qualified ranges

When the user types something like:
- `=SUM(Sheet2!A`
- `=SUM('My Sheet'!A`

`suggestSchemaRanges()` will:
1. parse the `SheetName!` prefix
2. require the sheet name to match a real sheet name (case-insensitive)
3. require the user to have already typed quotes when the sheet name needs quoting (spaces, leading digits, names that look like A1/R1C1 refs, etc)
4. delegate to `suggestRanges()` with `sheetName` so `getCellValue()` can read the correct sheet

This is intentionally conservative: we avoid “fixing” missing quotes or incomplete sheet names because those would require edits *before* the caret and can’t be shown as ghost text in the formula bar.

---

## 3) Range suggestion heuristics (`rangeSuggester.js`)

Range suggestions are implemented in `packages/ai-completion/src/rangeSuggester.js` and are deliberately simple so they run inside the latency budget.

### Current algorithm (today)

`suggestRanges({ currentArgText, cellRef, surroundingCells, sheetName?, maxScanRows?, maxScanCols? })`:

1. **Accepts conservative A1-style prefixes + single-column partial range syntax (pure-insertion friendly)**
   - Base token regex (no `:`): `^(\$?)([A-Za-z]{1,3})(?:(\$?)(\d+))?$`
   - Supported prefix forms include:
     - `A`, `$A`, `A1`, `A$1`, `$A$1`
     - partial range prefixes: `A:`, `A1:`, `$A$1:`
     - single-column range prefixes: `A:A`, `A1:A` (and partially-typed end-column tokens for multi-letter columns, e.g. `AB1:A`)
   - Completion behavior for partial `:` forms:
     - `A1:` → can be extended to `A1:A10` (and `$A$1:` → `$A$1:$A$10`)
     - `A1:A` / `AB1:A` → end-column token is completed back to the start column (e.g. `AB1:A` → `AB1:AB10`)
     - `A:` → suggests `A:A` (entire column) as the pure-insertion-friendly completion
       - We intentionally do *not* suggest `A1:A10` here because that would require inserting characters *before* the typed `:`.
     - `A:A` is accepted as input (even if higher-confidence “contiguous block” suggestions may not always be pure-insertion extensions).
   - Empty argument (`currentArgText === ""`) is allowed: the suggester defaults to the **active cell’s column** so it can still propose ranges while remaining a pure insertion.
   - Still rejected/unsupported as inputs: multi-column prefixes like `A1:B` / `A1:B10`, structured refs like `Table1[Col]`, etc.
   - The suggester avoids “repairs” that would require inserting characters *before* the caret (e.g. it won’t turn `A:` into `A1:A10`).
     - For “partial” inputs (like `A`, `A1:`, `A1:A`), completions are designed to preserve the user-typed prefix so they can be rendered as a pure insertion in UIs that require it.
     - For already-formed-but-ambiguous inputs (like `A:A`), it may still emit alternative ranges (e.g. a contiguous block) that are **not** guaranteed to be pure insertions; UI surfaces may filter these out.

2. **Find a contiguous non-empty block in the typed column**
   - If the user provided a row (`A5`), treat that cell as the start and scan **down** until the first empty cell (`reason: contiguous_down_from_start`).
      - If the explicitly provided start cell is empty, no contiguous-block suggestion is produced.
   - Otherwise (user typed only a column, like `A`):
      - First scan **up** from the row above the current cell, skipping blank separators to find the nearest non-empty cell, then expand to the full contiguous block (`reason: contiguous_above_current_cell`).
      - **Downward scan fallback:** if there is no data above (within the scan budget), scan **down** using the same “skip blanks then expand” logic (`reason: contiguous_below_current_cell`).
        - If the active cell is in the referenced column, scanning starts at the row **below** the active cell.
        - If the active cell is in a different column (e.g. formula in `B2` referencing `A`), scanning starts on the **same row** so same-row values can be included.

3. **Numeric trimming heuristic (implicit scans)**
   - If a block is “mostly numeric”, non-numeric *edge* cells are treated as header/footer and trimmed (common for a text header row above numeric data).
   - This trimming is only applied when inferring a block above/below the current cell (not when the user explicitly types a start row like `A5`).

4. **Optional 2D table range suggestion (expand to the right)**
   - When a contiguous vertical block is found, the suggester also tries to expand that block **to the right** into a rectangular 2D range (e.g. `A1:D10`).
      - When available, the table detector uses the *untrimmed* block (before numeric header trimming) so the resulting 2D range can include a header row.
   - It grows column-by-column until it hits:
     - an entirely empty column (a “gap”), or
     - a column whose non-empty coverage over the rows is too low (currently `< 0.6`)
   - This is bounded by `maxScanCols` and returned with `reason` strings like `contiguous_table_above_current_cell` / `contiguous_table_below_current_cell` / `contiguous_table_down_from_start` (the table range is appended after the “entire column” suggestion for compatibility).

5. **Return up to three candidates**
   - `A1:A10` (contiguous 1D block, higher confidence)
   - `A:A` (entire column, lower confidence; always included)
   - optional `A1:D10` (2D table block)

### Hard bounds / caps

- `maxScanRows` (default **500**) caps vertical scanning (and therefore the number of rows inspected for table detection).
- `maxScanCols` (default **50**) caps rightward table expansion.

### Limits / known gaps

This is not a full “current region” detector. In particular, it:
- only expands **to the right** from the typed column (it does not scan left to find a table’s true start column)
- doesn’t infer row-only/horizontal ranges
- doesn’t look at formatting/table borders
- is intentionally conservative about complex/ambiguous references:
  - it rejects multi-column range prefixes like `A1:B` rather than guessing a 2D region
  - it only supports the limited prefix forms described above and won’t rewrite/repair arbitrary range syntax (including inserting missing characters before the caret)

### Adding new heuristics safely (no OOM, stay <100ms)

The biggest risks when extending range suggestions are:
- **OOM** from materializing large arrays/grids
- **latency regressions** from unbounded `getCellValue()` calls

Guidelines:
- **Always bound work explicitly.** Add caps like `maxScanRows`, `maxScanCols`, or a `maxCellReads` budget and enforce them.
- **Avoid building big intermediate structures.** Prefer streaming counts/metrics over collecting all values into arrays.
- **Design for “cheap fail”.** If the input is complex or ambiguous, return `[]` quickly rather than trying to be clever.
- **Keep suggestion lists tiny.** The engine only needs a handful of high-confidence candidates; return a small set with clear `confidence` and a `reason` string (the current module already does this).
- **Don’t force sheet materialization.** In desktop, cell reads can create sheets lazily; the integration layer intentionally blocks reads for unknown sheets. Preserve that pattern when adding multi-sheet logic.

If you add a new heuristic, add tests in `packages/ai-completion/test/rangeSuggester.test.js` (and/or `tabCompletionEngine.test.js`) that cover:
- scan limits being respected
- large/empty sheets returning quickly
- the new heuristic not allocating large memory

---

## 4) Desktop UI integration expectations (ghost text vs full replacement)

The completion engine returns **full draft strings** (`suggestion.text`). However, the desktop formula bar only displays suggestions as **ghost text** when they can be represented as a **pure insertion at the caret**.

### “Pure insertion” constraint

Ghost text is only renderable when:
- the suggestion starts with the exact prefix the user already typed (up to the cursor), and
- the suggestion preserves the exact suffix after the cursor (if any)

In other words: the suggestion must be “insert some characters at the caret” rather than “rewrite earlier characters”.

### Where this is enforced

`apps/desktop/src/ai/completion/formulaBarTabCompletion.ts`:
- requests suggestions via `TabCompletionEngine.getSuggestions()`
- chooses the first suggestion that passes `bestPureInsertionSuggestion(...)`
- calls `formulaBar.setAiSuggestion({ text, preview })`

`apps/desktop/src/formula-bar/FormulaBarModel.ts` then derives:
- `aiGhostText()` — the *inserted tail* used for rendering
- `acceptAiSuggestion()` — invoked on the **Tab** key to apply the full suggestion

### Why some suggestions intentionally don’t exist

Some completions are deliberately *not offered* because they would require edits before the caret and would never be renderable as ghost text. Examples (covered by tests):
- Completing `=SUM(My Sheet!A` by inserting the leading `'` quote is not offered (it would modify text before the cursor).
- Completing `=SUM(Tab[` to `=SUM(Table1[Amount])` is not offered (it would need to insert missing characters before the `[`).

When extending completions, assume that:
- suggestions that aren’t “pure insertion” may still be useful in other UIs,
- but the desktop formula bar ghost-text surface will ignore them.

---

## 5) Testing guidance

### Core engine + heuristics tests

Location:
- `packages/ai-completion/test/*.test.js` (Node’s built-in `node:test`)

Run:
- `pnpm test:node ai-completion`
  - (filters node:test files by substring; see `scripts/run-node-tests.mjs`)

Useful files:
- `packages/ai-completion/test/tabCompletionEngine.test.js`
- `packages/ai-completion/test/rangeSuggester.test.js`
- `packages/ai-completion/test/cursorTabCompletionClient.test.js`

### Desktop formula bar integration tests

Location:
- `apps/desktop/src/formula-bar/completion/tabCompletion.test.ts`
- `apps/desktop/src/formula-bar/completion/formulaBarTabCompletion.view.*`

Run:
- `pnpm -C apps/desktop vitest run src/formula-bar/completion/tabCompletion.test.ts`

Notes for test authors:
- keep tests network-safe and deterministic by stubbing/injecting a `completionClient` when needed (the engine works fine with `completionClient: null`).
