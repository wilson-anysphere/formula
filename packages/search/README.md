# `@formula/search`

Implementation of Excel-style **Find / Replace**, **Go To**, and **Name Box** parsing.

The API is designed to be:
- **stateful** (Find Next / Find Previous sessions);
- **cancellable** (AbortSignal);
- **scalable** (optional per-workbook indexing to avoid O(N) scans for repeated queries).

## Features

- Scopes: `selection`, `sheet`, `workbook`
- Look in: `values` (`display` vs `raw`) and `formulas`
- Options: match case, match entire cell contents
- Excel wildcards: `*`, `?`, `~` escape
- Search order: by rows / by columns
- Replace: next / all, across selection/sheet/workbook
- Merged cells: treated as the top-left ("master") cell, matching Excel address semantics
- Go To parser: A1 refs, sheet-qualified refs, named ranges, and table structured refs (including selectors like `#All`, `#Headers`, `#Data`, `#Totals`)

## API (high level)

```js
const {
  findAll,
  findNext,
  findPrev,
  replaceAll,
  replaceNext,
  parseGoTo,
  SearchSession,
  WorkbookSearchIndex,
} = require("./packages/search");
```

### SearchSession (recommended)

`SearchSession` caches matches and tracks the current position, so repeated `findNext` / `findPrev`
calls stay fast (and can wrap like Excel).

```js
const index = new WorkbookSearchIndex(workbook);

const session = new SearchSession(workbook, "needle", {
  scope: "workbook",
  currentSheetName: "Sheet1",
  from: { sheetName: "Sheet1", row: 0, col: 0 }, // optional starting cursor (exclusive)
  lookIn: "values",
  valueMode: "display",

  // Indexing (optional)
  index,
  indexStrategy: "auto", // "auto" | "always" | "never"

  // Cancellation / responsiveness
  signal: abortController.signal,
});

await session.findNext(); // => { sheetName, row, col, address, text, wrapped }
await session.findNext({ signal: abortController.signal }); // per-call override (optional)
await session.findPrev();
await session.replaceNext("replacement");
```

### Search options

```js
{
  scope: "selection" | "sheet" | "workbook",
  currentSheetName: "Sheet1",
  selectionRanges: [{ startRow, endRow, startCol, endCol }],
  from: { sheetName, row, col }, // optional (SearchSession only)

  lookIn: "values" | "formulas",
  valueMode: "display" | "raw", // when lookIn = "values"

  matchCase: boolean,
  matchEntireCell: boolean,
  useWildcards: boolean,
  searchOrder: "byRows" | "byColumns",
  wrap: boolean,

  // performance / cancellation
  // (internally yields based on elapsed time, not iteration counts)
  timeBudgetMs: number, // default: ~10ms
  checkEvery: number, // clock sampling frequency (default: 256)
  scheduler: () => Promise<void>, // optional custom yield mechanism
  signal: AbortSignal,

  // indexing
  index: WorkbookSearchIndex,
  indexStrategy: "auto" | "always" | "never",
}
```

## Notes on Excel semantics

Replacement behavior matches a common Excel workflow:

- **Look in: formulas** edits the cell input (formula string for formula cells; raw constant for value-only cells)
- **Look in: values** replaces the displayed/evaluated value and **overwrites formulas with constants**

This is covered by `packages/search/test/replace.test.js`.

### Merged cells

To support Excel-like merged-cell semantics, a sheet adapter can optionally implement:

- `getMergedRanges(): Array<{startRow,endRow,startCol,endCol}>`
- `getMergedMasterCell(row, col): {row,col} | null`

If present, selection scopes are expanded to include merged regions, and matches are always reported
at the merged cell's top-left address.
