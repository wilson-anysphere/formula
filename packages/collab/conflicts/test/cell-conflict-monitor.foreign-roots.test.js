import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

import * as Y from "yjs";

import { CellConflictMonitor } from "../src/cell-conflict-monitor.js";

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

test("CellConflictMonitor initializes when cells root was created by a different Yjs instance (CJS getMap)", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance eagerly instantiating the cells root.
  Ycjs.Doc.prototype.getMap.call(doc, "cells");
  assert.throws(() => doc.getMap("cells"), /different constructor/);

  const origin = { type: "local" };
  const monitor = new CellConflictMonitor({
    doc,
    localUserId: "user-a",
    origin,
    localOrigins: new Set([origin]),
    onConflict: () => {},
  });

  // Root should be normalized back to this module's Yjs constructors so `new Y.Map()`
  // cells can be inserted safely.
  assert.ok(doc.share.get("cells") instanceof Y.Map);
  assert.ok(doc.getMap("cells") instanceof Y.Map);

  monitor.setLocalValue("Sheet1:0:0", "hello");
  const cell = doc.getMap("cells").get("Sheet1:0:0");
  assert.ok(cell instanceof Y.Map);
  assert.equal(cell.get("value"), "hello");

  monitor.dispose();
  doc.destroy();
});

