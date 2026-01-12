import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold } from "../toolbar.js";

test("toolbar helpers block huge full-width row selections (band row cap)", () => {
  const doc = new DocumentController();

  // Full-width rows: 100,001 rows is too large to format (UI guard aligns with DocumentController
  // maxEnumeratedRows=50k). The toolbar helper should return false *without* calling into the
  // controller to avoid heavy work/delta allocations.
  let setRangeFormatCalls = 0;
  const originalSetRangeFormat = doc.setRangeFormat.bind(doc);
  doc.setRangeFormat = (...args) => {
    setRangeFormatCalls += 1;
    return originalSetRangeFormat(...args);
  };

  const applied = toggleBold(doc, "Sheet1", "A1:XFD100001", { next: true });
  assert.equal(applied, false);
  assert.equal(setRangeFormatCalls, 0);
});
