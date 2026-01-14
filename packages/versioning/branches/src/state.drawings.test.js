import assert from "node:assert/strict";
import test from "node:test";

import { normalizeDocumentState } from "./state.js";

test("normalizeDocumentState: filters drawings with oversized/invalid ids and trims string ids", () => {
  const oversized = "x".repeat(5000);

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
            drawings: [
              { id: oversized, zOrder: 0 },
              { id: "  ok  ", zOrder: 1 },
              { id: 1, zOrder: 2 },
              { id: "   ", zOrder: 3 },
              { id: Number.NaN, zOrder: 4 },
              { id: 2.5, zOrder: 5 },
              { id: null, zOrder: 6 },
            ],
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
  const view = normalized.sheets.metaById.Sheet1.view;
  assert.ok(view && typeof view === "object");
  assert.ok(Object.prototype.hasOwnProperty.call(view, "drawings"));
  assert.deepEqual(
    view.drawings.map((d) => d.id),
    ["ok", 1],
  );
});

test("normalizeDocumentState: preserves explicit drawings key even if all entries are dropped", () => {
  const oversized = "x".repeat(5000);

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
            drawings: [{ id: oversized }],
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
  const view = normalized.sheets.metaById.Sheet1.view;
  assert.ok(view && typeof view === "object");
  assert.ok(Object.prototype.hasOwnProperty.call(view, "drawings"));
  assert.deepEqual(view.drawings, []);
});

