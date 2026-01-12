import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold } from "../toolbar.js";

test("toolbar helpers refuse oversized full-width row selections (row-band guard) before calling setRangeFormat", () => {
  const doc = new DocumentController();
  let setRangeFormatCalls = 0;
  const originalSetRangeFormat = doc.setRangeFormat.bind(doc);
  doc.setRangeFormat = (...args) => {
    setRangeFormatCalls += 1;
    return originalSetRangeFormat(...args);
  };

  // Full-width rows: 100,001 rows exceeds the toolbar's row-band cap (50,000).
  // This is blocked *before* attempting to apply formatting via DocumentController.setRangeFormat.
  const applied = toggleBold(doc, "Sheet1", "A1:XFD100001", { next: true });
  assert.equal(applied, false);
  assert.equal(setRangeFormatCalls, 0);
});
