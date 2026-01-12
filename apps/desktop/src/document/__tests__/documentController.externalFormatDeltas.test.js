import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("applyExternalFormatDeltas updates getCellFormat, emits change.formatDeltas, and does not create undo history", () => {
  const doc = new DocumentController();

  // Ensure the sheet exists (DocumentController is lazily sheet-creating).
  assert.deepEqual(doc.getCellFormat("Sheet1", "A1"), {});

  const beforeDepth = doc.getStackDepths();

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
  assert.ok(lastChange, "expected a change event");
  assert.deepEqual(lastChange.formatDeltas, [{ sheetId: "Sheet1", layer: "sheet", beforeStyleId: 0, afterStyleId: boldId }]);
});

