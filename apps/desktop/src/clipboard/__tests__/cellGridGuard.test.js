import assert from "node:assert/strict";
import test from "node:test";

import { getCellGridFromRange } from "../clipboard.js";

test("getCellGridFromRange rejects huge ranges before scanning cells", () => {
  let scanned = 0;
  const doc = {
    getCell() {
      scanned += 1;
      throw new Error("Should not scan");
    },
  };

  assert.throws(() => getCellGridFromRange(doc, "Sheet1", "A1:Z8000"), /Range too large/i);
  assert.equal(scanned, 0);
});

