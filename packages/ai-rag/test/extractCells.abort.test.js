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
  // Fake a signal that becomes aborted after a few `.aborted` reads.
  // This lets us validate that `extractCells` checks abort status inside the
  // inner (cell) loop, not just once per row.
  let checks = 0;
  const signal = {
    get aborted() {
      checks += 1;
      return checks >= 4;
    },
  };

  const sheet = { cells: [[]] };
  const rect = { r0: 0, c0: 0, r1: 0, c1: 1000 };

  assert.throws(() => extractCells(sheet, rect, { signal }), { name: "AbortError" });
});

