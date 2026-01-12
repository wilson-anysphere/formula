import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("applyExternalFormatDeltas updates getCellFormat, emits change.formatDeltas, and does not create undo history", () => {
  const doc = new DocumentController();

  // Ensure the sheet exists (DocumentController is lazily sheet-creating).
  assert.deepEqual(doc.getCellFormat("Sheet1", "A1"), {});

  const beforeDepth = doc.getStackDepths();
  assert.equal(doc.isDirty, false);

  /** @type {any | null} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  const boldId = doc.styleTable.intern({ font: { bold: true } });
  doc.applyExternalFormatDeltas(
    [{ sheetId: "Sheet1", layer: "sheet", beforeStyleId: 0, afterStyleId: boldId }],
    { source: "collab" },
  );

  assert.deepEqual(doc.getCellFormat("Sheet1", "A1"), { font: { bold: true } });
  assert.deepEqual(doc.getStackDepths(), beforeDepth);
  assert.equal(doc.isDirty, true, "external deltas should mark the document dirty by default");
  assert.ok(lastChange, "expected a change event");
  assert.deepEqual(lastChange.formatDeltas, [{ sheetId: "Sheet1", layer: "sheet", beforeStyleId: 0, afterStyleId: boldId }]);
});

test("applyExternalFormatDeltas respects markDirty=false", () => {
  const doc = new DocumentController();
  doc.markSaved();
  assert.equal(doc.isDirty, false);

  const boldId = doc.styleTable.intern({ font: { bold: true } });
  doc.applyExternalFormatDeltas(
    [{ sheetId: "Sheet1", layer: "sheet", beforeStyleId: 0, afterStyleId: boldId }],
    { source: "collab", markDirty: false },
  );

  assert.deepEqual(doc.getCellFormat("Sheet1", "A1"), { font: { bold: true } });
  assert.equal(doc.isDirty, false);
});

test("applyExternalFormatDeltas bypasses canEditCell guards", () => {
  const doc = new DocumentController({
    canEditCell: () => {
      throw new Error("canEditCell should not be consulted for external deltas");
    },
  });

  const boldId = doc.styleTable.intern({ font: { bold: true } });
  doc.applyExternalFormatDeltas([{ sheetId: "Sheet1", layer: "sheet", beforeStyleId: 0, afterStyleId: boldId }], {
    source: "collab",
  });

  assert.deepEqual(doc.getCellFormat("Sheet1", "A1"), { font: { bold: true } });
});
