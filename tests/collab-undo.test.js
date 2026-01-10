import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { createUndoService, REMOTE_ORIGIN } from "../packages/collab/undo/index.js";

/**
 * @param {Y.Doc} docA
 * @param {Y.Doc} docB
 */
function connectDocs(docA, docB) {
  const forwardA = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docB, update, REMOTE_ORIGIN);
  };
  const forwardB = (update, origin) => {
    if (origin === REMOTE_ORIGIN) return;
    Y.applyUpdate(docA, update, REMOTE_ORIGIN);
  };

  docA.on("update", forwardA);
  docB.on("update", forwardB);

  // Initial bidirectional sync.
  Y.applyUpdate(docA, Y.encodeStateAsUpdate(docB), REMOTE_ORIGIN);
  Y.applyUpdate(docB, Y.encodeStateAsUpdate(docA), REMOTE_ORIGIN);

  return {
    disconnect() {
      docA.off("update", forwardA);
      docB.off("update", forwardB);
    }
  };
}

/**
 * @param {Y.Map<any>} cells
 * @param {string} cellKey
 */
function getFormula(cells, cellKey) {
  const cell = /** @type {Y.Map<any>|undefined} */ (cells.get(cellKey));
  return (cell?.get("formula") ?? "").toString();
}

/**
 * @param {object} opts
 * @param {Y.Map<any>} opts.cells
 * @param {ReturnType<typeof createUndoService>} opts.undo
 * @param {string} opts.cellKey
 * @param {string} opts.formula
 * @param {string} opts.userId
 */
function setFormula(opts) {
  opts.undo.perform({
    redo: () => {
      let cell = /** @type {Y.Map<any>|undefined} */ (opts.cells.get(opts.cellKey));
      if (!cell) {
        cell = new Y.Map();
        opts.cells.set(opts.cellKey, cell);
      }

      const nextFormula = opts.formula.trim();
      if (nextFormula) {
        cell.set("formula", nextFormula);
      } else {
        cell.delete("formula");
      }
      cell.set("modified", Date.now());
      cell.set("modifiedBy", opts.userId);
    }
  });
}

test("collab undo/redo only affects local changes", () => {
  const doc1 = new Y.Doc();
  const doc2 = new Y.Doc();
  const cells1 = doc1.getMap("cells");
  const cells2 = doc2.getMap("cells");

  const undo1 = createUndoService({ mode: "collab", doc: doc1, scope: cells1 });
  const undo2 = createUndoService({ mode: "collab", doc: doc2, scope: cells2 });

  connectDocs(doc1, doc2);

  setFormula({ cells: cells1, undo: undo1, cellKey: "sheet:0:0", formula: "=1", userId: "u1" });
  setFormula({ cells: cells2, undo: undo2, cellKey: "sheet:0:1", formula: "=2", userId: "u2" });

  assert.equal(getFormula(cells1, "sheet:0:0"), "=1");
  assert.equal(getFormula(cells1, "sheet:0:1"), "=2");
  assert.equal(getFormula(cells2, "sheet:0:0"), "=1");
  assert.equal(getFormula(cells2, "sheet:0:1"), "=2");

  undo1.undo();

  // Local A1 should undo, remote B1 must stay.
  assert.equal(getFormula(cells1, "sheet:0:0"), "");
  assert.equal(getFormula(cells1, "sheet:0:1"), "=2");
  assert.equal(getFormula(cells2, "sheet:0:0"), "");
  assert.equal(getFormula(cells2, "sheet:0:1"), "=2");

  undo1.redo();
  assert.equal(getFormula(cells1, "sheet:0:0"), "=1");
  assert.equal(getFormula(cells1, "sheet:0:1"), "=2");
  assert.equal(getFormula(cells2, "sheet:0:0"), "=1");
  assert.equal(getFormula(cells2, "sheet:0:1"), "=2");
});

test("collab undo batches rapid typing into a single undo step", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  const undo = createUndoService({ mode: "collab", doc, scope: cells, captureTimeoutMs: 10_000 });

  setFormula({ cells, undo, cellKey: "sheet:0:0", formula: "=1", userId: "u1" });
  setFormula({ cells, undo, cellKey: "sheet:0:0", formula: "=1+1", userId: "u1" });
  setFormula({ cells, undo, cellKey: "sheet:0:0", formula: "=1+1+1", userId: "u1" });

  assert.equal(getFormula(cells, "sheet:0:0"), "=1+1+1");

  undo.undo();

  // If the edits were captured as a single step, undo should return to empty.
  assert.equal(getFormula(cells, "sheet:0:0"), "");
});
