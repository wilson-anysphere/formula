import test from "node:test";
import assert from "node:assert/strict";

import { diffDocumentWorkbookSnapshots } from "../packages/versioning/src/index.js";

const encoder = new TextEncoder();

/**
 * @param {any} value
 */
function encodeSnapshot(value) {
  return encoder.encode(JSON.stringify(value));
}

test("diffDocumentWorkbookSnapshots reports workbook-level metadata changes (JSON snapshots)", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: [{ row: 0, col: 0, value: "move-me", formula: "=A1+B1", format: null }],
      },
      { id: "sheet2", name: "Sheet2", cells: [] },
    ],
    comments: {
      c1: { id: "c1", cellRef: "A1", content: "Original comment", resolved: false, replies: [] },
    },
    metadata: {
      title: "Budget",
      owner: "u1",
    },
    namedRanges: {
      NR1: { sheetId: "sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
    },
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Renamed",
        cells: [{ row: 2, col: 3, value: "move-me", formula: "=B1 + A1", format: null }],
      },
      { id: "sheet3", name: "Sheet3", cells: [] },
    ],
    comments: {
      c1: {
        id: "c1",
        cellRef: "A1",
        content: "Updated comment",
        resolved: true,
        replies: [{ id: "r1", content: "First reply" }],
      },
    },
    metadata: {
      title: "Budget (edited)",
      theme: { name: "dark" },
    },
    namedRanges: {
      NR1: { sheetId: "sheet1", rect: { r0: 0, c0: 0, r1: 3, c1: 3 } },
      NR2: { sheetId: "sheet1", rect: { r0: 1, c0: 1, r1: 2, c1: 2 } },
    },
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });

  assert.deepEqual(diff.sheets.renamed, [{ id: "sheet1", beforeName: "Sheet1", afterName: "Renamed" }]);
  assert.deepEqual(diff.sheets.added, [{ id: "sheet3", name: "Sheet3", afterIndex: 1 }]);
  assert.deepEqual(diff.sheets.removed, [{ id: "sheet2", name: "Sheet2", beforeIndex: 1 }]);
  assert.deepEqual(diff.sheets.moved, []);

  assert.deepEqual(
    diff.cellsBySheet.map((entry) => entry.sheetId),
    ["sheet1", "sheet2", "sheet3"],
  );
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.moved.length, 1);
  assert.deepEqual(sheet1Diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(sheet1Diff.moved[0].newLocation, { row: 2, col: 3 });

  assert.equal(diff.comments.modified.length, 1);
  assert.equal(diff.comments.modified[0].id, "c1");
  assert.equal(diff.comments.modified[0].after.repliesLength, 1);

  assert.deepEqual(diff.namedRanges.added.map((r) => r.key), ["NR2"]);
  assert.equal(diff.namedRanges.modified.length, 1);
  assert.equal(diff.namedRanges.modified[0].key, "NR1");

  assert.deepEqual(diff.metadata.added.map((r) => r.key), ["theme"]);
  assert.deepEqual(diff.metadata.removed.map((r) => r.key), ["owner"]);
  assert.equal(diff.metadata.modified.length, 1);
  assert.equal(diff.metadata.modified[0].key, "title");
});
