import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";
import * as Y from "yjs";

import { workbookStateFromYjsDoc } from "./workbookState.js";
import { sheetStateFromYjsDoc } from "./sheetState.js";

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

test("sheetStateFromYjsDoc tolerates foreign placeholder roots created via CJS Doc.get", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  // Simulate another Yjs module instance calling `Doc.get(name)` (defaulting to
  // AbstractType) on this doc, leaving a foreign placeholder constructor.
  Ycjs.Doc.prototype.get.call(doc, "cells");
  Ycjs.Doc.prototype.get.call(doc, "sheets");

  // Hydrate some content via a CJS update.
  const remote = new Ycjs.Doc();
  remote.getArray("sheets").push([{ id: "Sheet1", name: "Sheet1" }]);
  const cells = remote.getMap("cells");
  const cell = new Ycjs.Map();
  cell.set("value", "alpha");
  cell.set("formula", null);
  cells.set("Sheet1:0:0", cell);
  const update = Ycjs.encodeStateAsUpdate(remote);
  Ycjs.applyUpdate(doc, update);

  // Regression: local `doc.getMap/getArray` would throw pre-normalization.
  assert.throws(() => doc.getMap("cells"), /different constructor/);
  assert.throws(() => doc.getArray("sheets"), /different constructor/);

  const state = sheetStateFromYjsDoc(doc, { sheetId: "Sheet1" });
  assert.equal(state.cells.get("r0c0")?.value ?? null, "alpha");

  // Root placeholders should be normalized so local callers can use getMap/getArray.
  assert.ok(doc.getMap("cells") instanceof Y.Map);
  assert.ok(doc.getArray("sheets") instanceof Y.Array);

  doc.destroy();
  remote.destroy();
});

test("workbookStateFromYjsDoc tolerates foreign placeholder roots created via CJS Doc.get", () => {
  const Ycjs = requireYjsCjs();

  const doc = new Y.Doc();

  Ycjs.Doc.prototype.get.call(doc, "cells");
  Ycjs.Doc.prototype.get.call(doc, "sheets");

  const remote = new Ycjs.Doc();
  remote.getArray("sheets").push([{ id: "Sheet1", name: "Sheet1" }]);
  const cells = remote.getMap("cells");
  const cell = new Ycjs.Map();
  cell.set("value", "alpha");
  cell.set("formula", null);
  cells.set("Sheet1:0:0", cell);
  const update = Ycjs.encodeStateAsUpdate(remote);
  Ycjs.applyUpdate(doc, update);

  assert.throws(() => doc.getMap("cells"), /different constructor/);
  assert.throws(() => doc.getArray("sheets"), /different constructor/);

  const state = workbookStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets, [
    { id: "Sheet1", name: "Sheet1", visibility: "visible", tabColor: null, view: { frozenRows: 0, frozenCols: 0 } },
  ]);
  assert.deepEqual(state.sheetOrder, ["Sheet1"]);
  assert.equal(state.cellsBySheet.get("Sheet1")?.cells.get("r0c0")?.value ?? null, "alpha");

  assert.ok(doc.getMap("cells") instanceof Y.Map);
  assert.ok(doc.getArray("sheets") instanceof Y.Array);

  doc.destroy();
  remote.destroy();
});
