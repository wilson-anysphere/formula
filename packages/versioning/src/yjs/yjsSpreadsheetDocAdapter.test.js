import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { createYjsSpreadsheetDocAdapter } from "./yjsSpreadsheetDocAdapter.js";

test('createYjsSpreadsheetDocAdapter.applyState uses origin "versioning-restore"', (t) => {
  const doc = new Y.Doc();
  t.after(() => doc.destroy());

  const adapter = createYjsSpreadsheetDocAdapter(doc);
  const cells = doc.getMap("cells");

  // Seed a simple workbook cell using the canonical collab cell schema (Y.Map).
  const cellA = new Y.Map();
  cellA.set("value", "alpha");
  cellA.set("formula", null);
  cells.set("Sheet1:0:0", cellA);

  const snapshot = adapter.encodeState();

  // Mutate the doc so applyState has real work to do.
  const cellB = new Y.Map();
  cellB.set("value", "beta");
  cellB.set("formula", null);
  cells.set("Sheet1:0:0", cellB);

  /** @type {any[]} */
  const origins = [];
  const onAfterTx = (tx) => origins.push(tx?.origin);
  doc.on("afterTransaction", onAfterTx);
  t.after(() => {
    doc.off("afterTransaction", onAfterTx);
  });

  adapter.applyState(snapshot);

  assert.deepEqual(origins, ["versioning-restore"]);

  const restored = /** @type {any} */ (cells.get("Sheet1:0:0"));
  assert.equal(restored?.get?.("value") ?? null, "alpha");
  assert.equal(restored?.get?.("formula") ?? null, null);
});

