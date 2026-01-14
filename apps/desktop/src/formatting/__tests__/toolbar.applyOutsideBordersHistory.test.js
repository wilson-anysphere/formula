import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { applyOutsideBorders } from "../toolbar.js";

test("applyOutsideBorders produces a single undo step when called directly", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [
    ["x", "y"],
    ["z", "w"],
  ]);

  const before = doc.history.length;
  const ok = applyOutsideBorders(doc, "Sheet1", "A1:B2", { style: "thin", color: "#FF000000" });
  assert.equal(ok, true);
  assert.equal(doc.history.length, before + 1);
});

test("applyOutsideBorders does not start a nested batch when already batching", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [
    ["x", "y"],
    ["z", "w"],
  ]);

  const historyBefore = doc.history.length;
  doc.beginBatch({ label: "Outer batch" });
  assert.equal(doc.batchDepth, 1);

  const ok = applyOutsideBorders(doc, "Sheet1", "A1:B2", { style: "thin", color: "#FF000000" });
  assert.equal(ok, true);
  assert.equal(doc.batchDepth, 1);

  doc.endBatch();
  assert.equal(doc.batchDepth, 0);
  assert.equal(doc.history.length, historyBefore + 1);
});
