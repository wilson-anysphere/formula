import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../apps/desktop/src/document/documentController.js";
import { diffDocumentWorkbookSnapshots } from "../packages/versioning/src/document/diffWorkbookSnapshots.js";

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
  assert.deepEqual(diff.sheets.added, [
    { id: "sheet3", name: "Sheet3", afterIndex: 1, visibility: "visible", tabColor: null, view: { frozenRows: 0, frozenCols: 0 } },
  ]);
  assert.deepEqual(diff.sheets.removed, [
    { id: "sheet2", name: "Sheet2", beforeIndex: 1, visibility: "visible", tabColor: null, view: { frozenRows: 0, frozenCols: 0 } },
  ]);
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

test("diffDocumentWorkbookSnapshots reports sheet metadata changes (visibility/tabColor/frozen panes)", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        visibility: "visible",
        tabColor: { rgb: "FF00FF00" },
        frozenRows: 1,
        frozenCols: 0,
        cells: [],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        visibility: "hidden",
        tabColor: null,
        frozenRows: 2,
        frozenCols: 3,
        cells: [],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });

  assert.deepEqual(diff.sheets.added, []);
  assert.deepEqual(diff.sheets.removed, []);
  assert.deepEqual(diff.sheets.renamed, []);
  assert.deepEqual(diff.sheets.moved, []);
  assert.deepEqual(diff.sheets.metaChanged, [
    { id: "sheet1", field: "tabColor", before: "FF00FF00", after: null },
    { id: "sheet1", field: "view.frozenCols", before: 0, after: 3 },
    { id: "sheet1", field: "view.frozenRows", before: 1, after: 2 },
    { id: "sheet1", field: "visibility", before: "visible", after: "hidden" },
  ]);
});

test("diffDocumentWorkbookSnapshots reads frozen panes from nested sheet.view", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        view: { frozenRows: 0, frozenCols: 1 },
        cells: [],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        view: { frozenRows: 4, frozenCols: 2 },
        cells: [],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.metaChanged, [
    { id: "sheet1", field: "view.frozenCols", before: 1, after: 2 },
    { id: "sheet1", field: "view.frozenRows", before: 0, after: 4 },
  ]);
});

test("diffDocumentWorkbookSnapshots canonicalizes tabColor to 8-digit ARGB", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        tabColor: { rgb: "#00FF00" },
        cells: [],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        tabColor: null,
        cells: [],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.metaChanged, [{ id: "sheet1", field: "tabColor", before: "FF00FF00", after: null }]);
});

test("diffDocumentWorkbookSnapshots accepts tabColor.argb (ExcelJS style) and canonicalizes to 8-digit ARGB", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        tabColor: { argb: "#00FF00" },
        cells: [],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        tabColor: null,
        cells: [],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.metaChanged, [{ id: "sheet1", field: "tabColor", before: "FF00FF00", after: null }]);
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits when default formats change (layered formats)", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        // A1 exists (non-empty) but has no per-cell format override.
        cells: [{ row: 0, col: 0, value: "x", formula: null, format: null }],
        defaultFormat: null,
        rowFormats: [],
        // Column A default formatting (matches DocumentController.encodeState output shape).
        colFormats: [{ col: 0, format: { font: { bold: true } } }],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: [{ row: 0, col: 0, value: "x", formula: null, format: null }],
        defaultFormat: null,
        rowFormats: [],
        // Same cell content, but the column default format changed.
        colFormats: [{ col: 0, format: { font: { italic: true } } }],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);

  assert.deepEqual(sheet1Diff.added, []);
  assert.deepEqual(sheet1Diff.removed, []);
  assert.deepEqual(sheet1Diff.modified, []);
  assert.deepEqual(sheet1Diff.moved, []);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffDocumentWorkbookSnapshots prefers explicit sheetOrder when present", () => {
  // The `sheets` array order is not authoritative when `sheetOrder` is present.
  // This matches newer DocumentController snapshots, which encode a redundant
  // `sheetOrder` field for robustness even if the `sheets` array is manipulated.
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheetOrder: ["sheet2", "sheet1"],
    // Intentionally not matching `sheetOrder`.
    sheets: [
      { id: "sheet1", name: "Sheet1", cells: [] },
      { id: "sheet2", name: "Sheet2", cells: [] },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheetOrder: ["sheet1", "sheet2"],
    sheets: [
      { id: "sheet1", name: "Sheet1", cells: [] },
      { id: "sheet2", name: "Sheet2", cells: [] },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.added, []);
  assert.deepEqual(diff.sheets.removed, []);
  assert.deepEqual(diff.sheets.renamed, []);
  assert.deepEqual(diff.sheets.moved, [{ id: "sheet2", beforeIndex: 0, afterIndex: 1 }]);
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits for column-default formatting (DocumentController snapshots)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");
  const beforeSnapshot = doc.encodeState();

  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });
  const afterSnapshot = doc.encodeState();

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1 = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1");
  assert.ok(sheet1);
  assert.equal(sheet1.diff.formatOnly.length, 1);
  assert.deepEqual(sheet1.diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits for row-default formatting (DocumentController snapshots)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");
  const beforeSnapshot = doc.encodeState();

  doc.setRangeFormat("Sheet1", "A1:XFD1", { font: { italic: true } });
  const afterSnapshot = doc.encodeState();

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1 = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1");
  assert.ok(sheet1);
  assert.equal(sheet1.diff.formatOnly.length, 1);
  assert.deepEqual(sheet1.diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits for sheet-default formatting (DocumentController snapshots)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");
  const beforeSnapshot = doc.encodeState();

  doc.setRangeFormat("Sheet1", "A1:XFD1048576", { font: { bold: true } });
  const afterSnapshot = doc.encodeState();

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1 = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1");
  assert.ok(sheet1);
  assert.equal(sheet1.diff.formatOnly.length, 1);
  assert.deepEqual(sheet1.diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits for formatRunsByCol range-run formatting (DocumentController snapshots)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", "x");
  const beforeSnapshot = doc.encodeState();

  // This range is:
  // - not full-sheet
  // - not full-height columns
  // - not full-width rows
  // - above RANGE_RUN_FORMAT_THRESHOLD (50k) so it should be encoded as `formatRunsByCol`.
  doc.setRangeFormat("Sheet1", "A1:Z2000", { font: { bold: true } });
  const afterSnapshot = doc.encodeState();

  const parsed = JSON.parse(new TextDecoder().decode(afterSnapshot));
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.ok(sheet);
  assert.ok(Array.isArray(sheet.formatRunsByCol) && sheet.formatRunsByCol.length > 0);
  assert.equal(
    sheet.cells.find((c) => c.row === 0 && c.col === 0)?.format ?? null,
    null,
    "A1 should not have a per-cell format override when encoded via range runs",
  );

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1 = diff.cellsBySheet.find((entry) => entry.sheetId === "Sheet1");
  assert.ok(sheet1);
  assert.equal(sheet1.diff.formatOnly.length, 1);
  assert.deepEqual(sheet1.diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits when range format runs change (layered formats)", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: [{ row: 0, col: 0, value: "x", formula: null, format: null }],
        defaultFormat: null,
        rowFormats: [],
        colFormats: [],
        // Optional Task 118-style sparse range formats (covers A1 only).
        formatRuns: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0, format: { font: { bold: true } } }],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: [{ row: 0, col: 0, value: "x", formula: null, format: null }],
        defaultFormat: null,
        rowFormats: [],
        colFormats: [],
        formatRuns: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 0, format: { font: { italic: true } } }],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);

  assert.deepEqual(sheet1Diff.added, []);
  assert.deepEqual(sheet1Diff.removed, []);
  assert.deepEqual(sheet1Diff.modified, []);
  assert.deepEqual(sheet1Diff.moved, []);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffDocumentWorkbookSnapshots reports formatOnly edits when sheet default format changes (layered formats)", () => {
  const beforeSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: [{ row: 0, col: 0, value: "x", formula: null, format: null }],
        defaultFormat: null,
        rowFormats: [],
        colFormats: [],
      },
    ],
  });

  const afterSnapshot = encodeSnapshot({
    schemaVersion: 1,
    sheets: [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: [{ row: 0, col: 0, value: "x", formula: null, format: null }],
        defaultFormat: { font: { bold: true } },
        rowFormats: [],
        colFormats: [],
      },
    ],
  });

  const diff = diffDocumentWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);

  assert.deepEqual(sheet1Diff.added, []);
  assert.deepEqual(sheet1Diff.removed, []);
  assert.deepEqual(sheet1Diff.modified, []);
  assert.deepEqual(sheet1Diff.moved, []);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});
