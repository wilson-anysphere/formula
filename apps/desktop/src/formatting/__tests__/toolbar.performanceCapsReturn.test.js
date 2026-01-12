import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold } from "../toolbar.js";

test("toolbar helpers propagate setRangeFormat cap skips via return value", () => {
  const doc = new DocumentController();

  const warnings = [];
  const originalWarn = console.warn;
  console.warn = (...args) => {
    warnings.push(args.join(" "));
  };

  try {
    // Full-width rows: 100,001 rows exceeds DocumentController's default maxEnumeratedRows (50,000),
    // so formatting is skipped and the toolbar helper should return false.
    const applied = toggleBold(doc, "Sheet1", "A1:XFD100001", { next: true });
    assert.equal(applied, false);
  } finally {
    console.warn = originalWarn;
  }

  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /Skipping row formatting/);
});

