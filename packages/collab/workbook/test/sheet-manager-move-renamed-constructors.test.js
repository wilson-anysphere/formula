import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { ensureWorkbookSchema, SheetManager } from "../src/index.ts";

test("SheetManager.moveSheet clones sheets even when their constructors are renamed", () => {
  const doc = new Y.Doc();
  const roots = ensureWorkbookSchema(doc);

  // Add a second sheet so we can exercise move semantics.
  doc.transact(() => {
    const YMapCtor = roots.cells.constructor;
    const sheet = new YMapCtor();
    sheet.set("id", "Sheet2");
    sheet.set("name", "Sheet2");
    sheet.set("visibility", "visible");
    roots.sheets.push([sheet]);
  });

  const mgr = new SheetManager({ doc });
  assert.deepEqual(mgr.list().map((s) => s.id), ["Sheet1", "Sheet2"]);

  const sheet1 = mgr.getById("Sheet1");
  assert.ok(sheet1);

  // Simulate a bundler-renamed constructor without mutating global `yjs` state.
  class RenamedMap extends sheet1.constructor {}
  Object.setPrototypeOf(sheet1, RenamedMap.prototype);

  mgr.moveSheet("Sheet1", 1);

  assert.deepEqual(mgr.list().map((s) => s.id), ["Sheet2", "Sheet1"]);
  const moved = mgr.getById("Sheet1");
  assert.ok(moved);
  assert.ok(moved instanceof Y.Map, "expected moved sheet entry to remain a Y.Map");
  assert.equal(moved.get("name"), "Sheet1");

  doc.destroy();
});

