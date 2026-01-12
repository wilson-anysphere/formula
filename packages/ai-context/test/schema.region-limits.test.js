import assert from "node:assert/strict";
import test from "node:test";

import { detectDataRegions } from "../src/schema.js";

test("detectDataRegions: caps scanned area to avoid huge visited allocations", () => {
  // 1,000 rows x 1,000 cols => 1,000,000 cells.
  // The region scanner should cap to 200,000 cells by truncating columns.
  const wideRow = Array.from({ length: 1_000 }, () => "x");
  const values = [
    wideRow,
    wideRow,
    // Remaining rows are empty but keep `rowCount` large enough to trigger the cap.
    ...Array.from({ length: 998 }, () => []),
  ];

  const regions = detectDataRegions(values);
  assert.ok(regions.length >= 1);
  assert.deepEqual(regions[0], { startRow: 0, startCol: 0, endRow: 1, endCol: 199 });
});

