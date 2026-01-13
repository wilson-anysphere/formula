import assert from "node:assert/strict";
import test from "node:test";

import { extractCells } from "../src/workbook/extractCells.js";

test("extractCells throws AbortError when signal already aborted", () => {
  const abortController = new AbortController();
  abortController.abort();

  const sheet = { cells: [[{ v: 1 }]] };
  const rect = { r0: 0, c0: 0, r1: 0, c1: 0 };

  assert.throws(() => extractCells(sheet, rect, { signal: abortController.signal }), {
    name: "AbortError",
  });
});

test("extractCells checks AbortSignal periodically during extraction", () => {
  const abortController = new AbortController();
  let cellReads = 0;
  const sheet = {
    getCell() {
      cellReads += 1;
      // Abort after extraction has started (but before the next periodic check).
      if (cellReads === 10) abortController.abort();
      return null;
    },
  };

  // Ensure we run enough cells to hit the periodic inner-loop abort check.
  const rect = { r0: 0, c0: 0, r1: 0, c1: 5000 };

  assert.throws(() => extractCells(sheet, rect, { signal: abortController.signal }), { name: "AbortError" });
});
