import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { applyDocumentStateToYjsDoc, yjsDocToDocumentState } from "./documentState.js";

test("yjs document adapter: normalizes formulas + handles legacy cell keys", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const cellA1 = new Y.Map();
  cellA1.set("value", 123);
  cells.set("Sheet1:0:0", cellA1);

  // Legacy `${sheetId}:${row},${col}` encoding + formula missing "=".
  const cellB1 = new Y.Map();
  cellB1.set("formula", "1+1");
  cellB1.set("format", { bold: true });
  cells.set("Sheet1:0,1", cellB1);

  // Unit-test convenience encoding.
  const cellC1 = new Y.Map();
  cellC1.set("formula", "=SUM(1,2)");
  cells.set("r0c2", cellC1);

  const state = yjsDocToDocumentState(doc);
  assert.deepEqual(state, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: { id: "Sheet1", name: "Sheet1" },
      },
    },
    cells: {
      Sheet1: {
        A1: { value: 123 },
        B1: { formula: "=1+1", format: { bold: true } },
        C1: { formula: "=SUM(1,2)" },
      },
    },
    namedRanges: {},
    comments: {},
  });

  const doc2 = new Y.Doc();
  applyDocumentStateToYjsDoc(doc2, state, { origin: { test: true } });

  const cells2 = doc2.getMap("cells");
  assert.equal(cells2.has("Sheet1:0:0"), true);
  assert.equal(cells2.has("Sheet1:0:1"), true);
  assert.equal(cells2.has("Sheet1:0:2"), true);
  assert.equal(cells2.has("Sheet1:0,1"), false);
  assert.equal(cells2.has("r0c2"), false);

  const b1 = /** @type {Y.Map<any>} */ (cells2.get("Sheet1:0:1"));
  assert.ok(b1);
  assert.equal(b1.get("formula"), "=1+1");
  assert.deepEqual(b1.get("format"), { bold: true });
});
