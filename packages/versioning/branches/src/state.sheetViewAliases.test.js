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
