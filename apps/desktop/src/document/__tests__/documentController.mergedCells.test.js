import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("mergeCells stores merged ranges in sheet view and clears non-anchor contents (preserving formats)", () => {
  const doc = new DocumentController();

  doc.setRangeValues("Sheet1", "A1", [
    [1, 2],
    [3, 4],
  ]);
  doc.setRangeFormat("Sheet1", "B2", { bold: true });

  const beforeB2 = doc.getCell("Sheet1", "B2");
  assert.notEqual(beforeB2.styleId, 0);

  doc.mergeCells("Sheet1", { startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);
  assert.deepEqual(doc.getMergedRangeAt("Sheet1", 1, 1), { startRow: 0, endRow: 1, startCol: 0, endCol: 1 });

  // Only the top-left cell retains its value.
  assert.equal(doc.getCell("Sheet1", "A1").value, 1);
  assert.equal(doc.getCell("Sheet1", "B1").value, null);
  assert.equal(doc.getCell("Sheet1", "A2").value, null);
  assert.equal(doc.getCell("Sheet1", "B2").value, null);

  // Formatting on cleared cells is preserved.
  const afterB2 = doc.getCell("Sheet1", "B2");
  assert.equal(afterB2.styleId, beforeB2.styleId);
});

test("mergeCells automatically removes overlapping merges (new merge wins)", () => {
  const doc = new DocumentController();
  doc.mergeCells("Sheet1", { startRow: 0, endRow: 0, startCol: 0, endCol: 1 }); // A1:B1
  doc.mergeCells("Sheet1", { startRow: 0, endRow: 1, startCol: 1, endCol: 2 }); // B1:C2 (overlaps)

  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 1, startCol: 1, endCol: 2 }]);
});

test("unmergeCells removes merges that intersect the target (cell or range)", () => {
  const doc = new DocumentController();
  doc.mergeCells("Sheet1", { startRow: 0, endRow: 1, startCol: 0, endCol: 1 }); // A1:B2
  doc.mergeCells("Sheet1", { startRow: 0, endRow: 1, startCol: 3, endCol: 4 }); // D1:E2

  // Unmerge by cell inside the first merge.
  doc.unmergeCells("Sheet1", { row: 1, col: 1 });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 1, startCol: 3, endCol: 4 }]);

  // Unmerge by range intersecting the remaining merge.
  doc.unmergeCells("Sheet1", { startRow: 0, endRow: 10, startCol: 0, endCol: 10 });
  assert.deepEqual(doc.getMergedRanges("Sheet1"), []);
});

test("encodeState/applyState roundtrip preserves merged cells", () => {
  const doc = new DocumentController();
  doc.mergeCells("Sheet1", { startRow: 2, endRow: 4, startCol: 1, endCol: 3 });

  const snapshot = doc.encodeState();
  const restored = new DocumentController();
  restored.applyState(snapshot);

  assert.deepEqual(restored.getMergedRanges("Sheet1"), doc.getMergedRanges("Sheet1"));
});

test("applyState accepts singleton-wrapped mergedRanges coordinates (interop)", () => {
  const snapshot = new TextEncoder().encode(
    JSON.stringify({
      schemaVersion: 1,
      sheets: [
        {
          id: "Sheet1",
          name: "Sheet1",
          visibility: "visible",
          frozenRows: 0,
          frozenCols: 0,
          cells: [],
          view: {
            mergedRanges: [
              { startRow: { 0: 0 }, endRow: [1], startCol: { 0: 0 }, endCol: [1] },
            ],
          },
        },
      ],
      sheetOrder: ["Sheet1"],
    }),
  );

  const doc = new DocumentController();
  doc.applyState(snapshot);
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);
});

test("applyState accepts singleton-wrapped mergedRanges entry objects (interop)", () => {
  const snapshot = new TextEncoder().encode(
    JSON.stringify({
      schemaVersion: 1,
      sheets: [
        {
          id: "Sheet1",
          name: "Sheet1",
          visibility: "visible",
          frozenRows: 0,
          frozenCols: 0,
          cells: [],
          view: {
            mergedRanges: [
              {
                0: { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
              },
            ],
          },
        },
      ],
      sheetOrder: ["Sheet1"],
    }),
  );

  const doc = new DocumentController();
  doc.applyState(snapshot);
  assert.deepEqual(doc.getMergedRanges("Sheet1"), [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);
});
