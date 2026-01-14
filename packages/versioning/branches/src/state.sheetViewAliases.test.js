import assert from "node:assert/strict";
import test from "node:test";
import * as Y from "yjs";

import { normalizeDocumentState } from "./state.js";

test("normalizeDocumentState: normalizes legacy backgroundImage/background_image to backgroundImageId", () => {
  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            backgroundImage: "  bg.png  ",
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    backgroundImageId: "bg.png",
  });
});

test("normalizeDocumentState: backgroundImageId=null overrides legacy aliases (prevents resurrection)", () => {
  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            backgroundImageId: null,
            background_image: "bg.png",
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    backgroundImageId: null,
  });
});

test("normalizeDocumentState: normalizes legacy merged range aliases to mergedRanges", () => {
  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            merged_ranges: [{ start_row: 0, end_row: 1, start_col: 0, end_col: 2 }],
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
  });
});

test("normalizeDocumentState: mergedRanges=null overrides legacy aliases (prevents resurrection)", () => {
  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            mergedRanges: null,
            merged_ranges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    mergedRanges: [],
  });
});

test("normalizeDocumentState: normalizes mergedRanges list items stored as Y.Maps (range wrapper + start/end objects)", () => {
  const doc = new Y.Doc();

  const yRanges = new Y.Array();
  const entry = new Y.Map();
  entry.set("range", {
    start: { row: 0, col: 0 },
    end: { row: 1, col: 2 },
  });
  yRanges.push([entry]);

  // Integrate the nested types into a doc so Yjs doesn't warn on reads.
  doc.transact(() => {
    doc.getArray("__root").push([yRanges]);
  });

  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            mergedRanges: yRanges,
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
  });

  doc.destroy();
});

test("normalizeDocumentState: normalizes axis overrides stored as Y.Maps and Y.Arrays", () => {
  const doc = new Y.Doc();

  const colWidthsMap = new Y.Map();
  colWidthsMap.set("0", 120);
  colWidthsMap.set("1", -5); // invalid
  colWidthsMap.set("2", 0); // invalid
  colWidthsMap.set("3", 45);

  const rowHeightsArr = new Y.Array();
  rowHeightsArr.push([
    [0, 10],
    [1, -1],
    { index: 2, size: 25 },
    (() => {
      const entry = new Y.Map();
      entry.set("index", 3);
      entry.set("size", 30);
      return entry;
    })(),
  ]);

  doc.transact(() => {
    doc.getArray("__root").push([colWidthsMap, rowHeightsArr]);
  });

  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            colWidths: colWidthsMap,
            rowHeights: rowHeightsArr,
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    colWidths: { "0": 120, "3": 45 },
    rowHeights: { "0": 10, "2": 25, "3": 30 },
  });

  doc.destroy();
});

test("normalizeDocumentState: normalizes format overrides stored as Y.Maps", () => {
  const doc = new Y.Doc();
  const rowFormats = new Y.Map();
  rowFormats.set("0", { font: { bold: true } });
  rowFormats.set("1", {}); // treated as empty/no-op format

  const runsByCol = new Y.Map();
  runsByCol.set("0", [{ startRow: 0, endRowExclusive: 1, format: { numberFormat: "0%" } }]);

  doc.transact(() => {
    doc.getArray("__root").push([rowFormats, runsByCol]);
  });

  const input = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            rowFormats,
            formatRunsByCol: runsByCol,
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  const normalized = normalizeDocumentState(input);
  assert.deepEqual(normalized.sheets.metaById.Sheet1.view, {
    frozenRows: 0,
    frozenCols: 0,
    rowFormats: { "0": { font: { bold: true } } },
    formatRunsByCol: [{ col: 0, runs: [{ startRow: 0, endRowExclusive: 1, format: { numberFormat: "0%" } }] }],
  });

  doc.destroy();
});
