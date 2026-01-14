import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

test("CellStructuralConflictMonitor records format-only edits without marking contentChanged", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  const origin = { type: "local" };

  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "local",
    origin,
    localOrigins: new Set([origin]),
    onConflict: () => {},
  });

  const cellKey = "Sheet1:0:0";

  // Seed a plaintext cell with a format payload.
  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "hello");
    cell.set("format", { font: { bold: true } });
    cells.set(cellKey, cell);
  }, origin);

  const ops = doc.getMap("cellStructuralOps");
  const beforeIds = new Set(Array.from(ops.keys(), (k) => String(k)));

  // Format-only update: should not be treated as contentChanged.
  doc.transact(() => {
    const cell = cells.get(cellKey);
    assert.ok(cell instanceof Y.Map);
    cell.set("format", { font: { bold: false } });
  }, origin);

  const afterIds = Array.from(ops.keys(), (k) => String(k));
  const created = afterIds.filter((id) => !beforeIds.has(id));
  assert.equal(created.length, 1, "expected exactly one op record for format-only edit");

  const record = ops.get(created[0]);
  assert.ok(record && typeof record === "object");
  assert.equal(record.kind, "edit");
  assert.equal(record.cellKey, cellKey);
  assert.equal(record.contentChanged, false);
  assert.equal(record.formatChanged, true);

  monitor.dispose();
});

