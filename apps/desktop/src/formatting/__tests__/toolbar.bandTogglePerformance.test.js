import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { toggleBold } from "../toolbar.js";

test("toggleBold on full-column band avoids per-cell format scans", () => {
  const doc = new DocumentController();

  let calls = 0;
  const original = doc.getCellFormat.bind(doc);
  doc.getCellFormat = (...args) => {
    calls += 1;
    return original(...args);
  };

  toggleBold(doc, "Sheet1", "A1:A1048576");

  // Restore before assertions that call `getCellFormat`.
  doc.getCellFormat = original;

  assert.ok(calls <= 10, `Expected <= 10 getCellFormat calls, got ${calls}`);

  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1")?.font?.bold), true);
  assert.equal(Boolean(doc.getCellFormat("Sheet1", "A1000")?.font?.bold), true);
});

