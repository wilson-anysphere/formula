import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { ensureWorkbookSchema, getWorkbookRoots } from "../src/index.ts";

test("ensureWorkbookSchema: filters oversized drawing ids when merging duplicate sheet view metadata", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversized = "x".repeat(5000);

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    loser.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      drawings: [{ id: oversized }, { id: " ok " }],
    });

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");
    winner.set("view", { frozenRows: 0, frozenCols: 0 });

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);

  const view = sheet.get("view");
  assert.ok(view && typeof view === "object");

  assert.ok(Array.isArray(view.drawings));
  assert.deepEqual(
    view.drawings.map((d) => d.id),
    ["ok"],
  );

  doc.destroy();
});

test("ensureWorkbookSchema: filters oversized drawing ids when copying a missing view payload from a duplicate sheet", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversized = "x".repeat(5000);

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    loser.set("view", {
      frozenRows: 0,
      frozenCols: 0,
      drawings: [{ id: oversized }, { id: "ok" }],
    });

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");
    // Winner intentionally lacks `view` so schema normalization has to copy it from the loser.

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);

  const view = sheet.get("view");
  assert.ok(view && typeof view === "object");
  assert.ok(Array.isArray(view.drawings));
  assert.deepEqual(
    view.drawings.map((d) => d.id),
    ["ok"],
  );

  doc.destroy();
});
