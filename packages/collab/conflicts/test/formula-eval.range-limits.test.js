import test from "node:test";
import assert from "node:assert/strict";

import { tryEvaluateFormula } from "../src/formula-eval.js";

test("tryEvaluateFormula rejects Excel-scale ranges before enumerating cells", () => {
  let scanned = 0;
  const result = tryEvaluateFormula("SUM(A1:Z8000)", {
    getCellValue() {
      scanned += 1;
      return 1;
    },
  });

  assert.equal(result.ok, false);
  assert.match(result.error, /range too large/i);
  assert.equal(scanned, 0);
});

test("tryEvaluateFormula MIN() handles large ranges without spread argument limits", () => {
  const result = tryEvaluateFormula("MIN(A1:CV1000)", {
    getCellValue({ row, col }) {
      // 0-based row/col; ensure the minimum value is at A1.
      return row * 100 + col;
    },
  });

  assert.equal(result.ok, true);
  if (!result.ok) throw new Error("Expected ok result");
  assert.equal(result.value, 0);
});

