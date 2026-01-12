import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { CellStructuralConflictMonitor } from "../src/cell-structural-conflict-monitor.js";

function requireYjsCjs() {
  const require = createRequire(import.meta.url);
  const prevError = console.error;
  console.error = (...args) => {
    if (typeof args[0] === "string" && args[0].startsWith("Yjs was already imported.")) return;
    prevError(...args);
  };
  try {
    // eslint-disable-next-line import/no-named-as-default-member
    return require("yjs");
  } finally {
    console.error = prevError;
  }
}

test("CellStructuralConflictMonitor preserves foreign cell maps even when constructors are renamed", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  const remoteCells = remote.getMap("cells");
  remote.transact(() => {
    const cell = new Ycjs.Map();
    cell.set("value", "from-cjs");
    cell.set("formula", null);
    cell.set("modified", 1);
    remoteCells.set("Sheet1:0:0", cell);
  });
  const update = Ycjs.encodeStateAsUpdate(remote);

  const doc = new Y.Doc();
  const cells = doc.getMap("cells");
  // Apply update via the CJS build to simulate y-websocket applying updates.
  Ycjs.applyUpdate(doc, update);

  const foreignCell = cells.get("Sheet1:0:0");
  assert.ok(foreignCell);
  assert.equal(foreignCell instanceof Y.Map, false);
  // Simulate a bundler-renamed constructor without mutating the global Yjs module state
  // (which can cause cross-test interference under concurrency).
  class RenamedMap extends foreignCell.constructor {}
  Object.setPrototypeOf(foreignCell, RenamedMap.prototype);

  const origin = { type: "local" };
  const monitor = new CellStructuralConflictMonitor({
    doc,
    cells,
    localUserId: "user-a",
    origin,
    localOrigins: new Set([origin]),
    onConflict: () => {},
  });

  monitor._clearCell("Sheet1:0:0");

  assert.equal(cells.get("Sheet1:0:0"), foreignCell, "expected monitor to mutate existing foreign map (no replacement)");
  assert.equal(foreignCell.get("value"), null);
  assert.equal(foreignCell.get("formula"), null);

  monitor.dispose();
  doc.destroy();
});
