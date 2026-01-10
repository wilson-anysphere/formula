# `@formula/search` (prototype)

Prototype implementation of Excel-style **Find / Replace**, **Go To**, and **Name Box** parsing.

This repository currently hosts mostly architecture docs; this package is a small, self-contained
implementation to unblock UI work and encode expected behavior in tests.

## Features

- Scopes: `selection`, `sheet`, `workbook`
- Look in: `values` (`display` vs `raw`) and `formulas`
- Options: match case, match entire cell contents
- Excel wildcards: `*`, `?`, `~` escape
- Search order: by rows / by columns
- Replace: next / all, across selection/sheet/workbook
- Go To parser: A1 refs, sheet-qualified refs, named ranges, minimal table structured refs

## API (high level)

```js
const {
  findAll,
  findNext,
  replaceAll,
  replaceNext,
  parseGoTo,
} = require("./packages/search");
```

### Search options

```js
{
  scope: "selection" | "sheet" | "workbook",
  currentSheetName: "Sheet1",
  selectionRanges: [{ startRow, endRow, startCol, endCol }],

  lookIn: "values" | "formulas",
  valueMode: "display" | "raw", // when lookIn = "values"

  matchCase: boolean,
  matchEntireCell: boolean,
  useWildcards: boolean,
  searchOrder: "byRows" | "byColumns",

  // performance / cancellation
  yieldEvery: number,
  signal: AbortSignal,
}
```

## Notes on Excel semantics

Replacement behavior matches a common Excel workflow:

- **Look in: formulas** edits the cell input (formula string for formula cells; raw constant for value-only cells)
- **Look in: values** replaces the displayed/evaluated value and **overwrites formulas with constants**

This is covered by `packages/search/test/replace.test.js`.

