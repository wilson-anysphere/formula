import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { ensureWorkbookSchema, SheetManager } from "../src/index.ts";

function createOversizedThrowingYText({ length = 5000 } = {}) {
  const text = new Y.Text();
  text.insert(0, "x".repeat(length));
  text.toString = () => {
    throw new Error("unexpected Y.Text.toString() on oversized drawing id");
  };
  text.toDelta = () => {
    throw new Error("unexpected Y.Text.toDelta() on oversized drawing id");
  };
  return text;
}

test("SheetManager.moveSheet does not clone oversized Y.Text drawing ids when reordering", () => {
  const doc = new Y.Doc();
  const roots = ensureWorkbookSchema(doc);

  // Add a second sheet so we can exercise move semantics.
  doc.transact(() => {
    const YMapCtor = roots.cells.constructor;
    const sheet2 = new YMapCtor();
    sheet2.set("id", "Sheet2");
    sheet2.set("name", "Sheet2");
    sheet2.set("visibility", "visible");
    roots.sheets.push([sheet2]);

    const sheet1 = roots.sheets.get(0);
    assert.ok(sheet1 instanceof Y.Map);

    const oversizedId = createOversizedThrowingYText();
    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);

    const drawings = new Y.Array();
    const bad = new Y.Map();
    bad.set("id", oversizedId);
    const ok = new Y.Map();
    ok.set("id", " ok ");
    drawings.push([bad, ok]);
    view.set("drawings", drawings);

    sheet1.set("view", view);
  });

  const mgr = new SheetManager({ doc });
  assert.deepEqual(mgr.list().map((s) => s.id), ["Sheet1", "Sheet2"]);

  // If moveSheet uses cloneYjsValue on the whole sheet map, it will attempt to clone the
  // oversized Y.Text id (calling toDelta/toString), and this test should fail.
  mgr.moveSheet("Sheet1", 1);

  assert.deepEqual(mgr.list().map((s) => s.id), ["Sheet2", "Sheet1"]);
  const moved = mgr.getById("Sheet1");
  assert.ok(moved);

  const view = moved.get("view");
  assert.ok(view && typeof view === "object");

  const drawings = view instanceof Y.Map ? view.get("drawings") : view.drawings;
  assert.ok(Array.isArray(drawings));
  assert.deepEqual(
    drawings.map((d) => d.id),
    ["ok"],
  );

  doc.destroy();
});

test("SheetManager.moveSheet does not clone oversized Y.Text drawing ids stored on top-level sheet.drawings", () => {
  const doc = new Y.Doc();
  const roots = ensureWorkbookSchema(doc);

  // Add a second sheet so we can exercise move semantics.
  doc.transact(() => {
    const YMapCtor = roots.cells.constructor;
    const sheet2 = new YMapCtor();
    sheet2.set("id", "Sheet2");
    sheet2.set("name", "Sheet2");
    sheet2.set("visibility", "visible");
    roots.sheets.push([sheet2]);

    const sheet1 = roots.sheets.get(0);
    assert.ok(sheet1 instanceof Y.Map);

    const oversizedId = createOversizedThrowingYText();
    const drawings = new Y.Array();
    const bad = new Y.Map();
    bad.set("id", oversizedId);
    const ok = new Y.Map();
    ok.set("id", " ok ");
    drawings.push([bad, ok]);

    sheet1.set("drawings", drawings);
  });

  const mgr = new SheetManager({ doc });
  assert.deepEqual(mgr.list().map((s) => s.id), ["Sheet1", "Sheet2"]);

  mgr.moveSheet("Sheet1", 1);

  assert.deepEqual(mgr.list().map((s) => s.id), ["Sheet2", "Sheet1"]);
  const moved = mgr.getById("Sheet1");
  assert.ok(moved);

  const drawings = moved.get("drawings");
  assert.ok(Array.isArray(drawings));
  assert.deepEqual(
    drawings.map((d) => d.id),
    ["ok"],
  );

  doc.destroy();
});
