import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { ensureWorkbookSchema, getWorkbookRoots } from "../src/index.ts";

function createOversizedThrowingYText({ length = 5000, throwOn = "toString" } = {}) {
  const text = new Y.Text();
  text.insert(0, "x".repeat(length));
  if (throwOn === "toString") {
    text.toString = () => {
      throw new Error("Unexpected call to Y.Text#toString");
    };
  }
  if (throwOn === "toDelta") {
    text.toDelta = () => {
      throw new Error("Unexpected call to Y.Text#toDelta");
    };
  }
  return text;
}

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

test("ensureWorkbookSchema: does not stringify oversized Y.Text drawing ids when merging mixed map/object view encodings", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversizedText = createOversizedThrowingYText({ throwOn: "toString" });

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");

    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);

    const drawings = new Y.Array();
    const bad = new Y.Map();
    bad.set("id", oversizedText);
    const ok = new Y.Map();
    ok.set("id", " ok ");
    drawings.push([bad, ok]);
    view.set("drawings", drawings);

    loser.set("view", view);

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

test("ensureWorkbookSchema: does not clone Y.Text drawing ids when copying a missing view payload from a duplicate sheet", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const oversizedText = createOversizedThrowingYText({ throwOn: "toDelta" });

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");

    const view = new Y.Map();
    view.set("frozenRows", 0);
    view.set("frozenCols", 0);

    const drawings = new Y.Array();
    const bad = new Y.Map();
    bad.set("id", oversizedText);
    const ok = new Y.Map();
    ok.set("id", "ok");
    drawings.push([bad, ok]);
    view.set("drawings", drawings);

    loser.set("view", view);

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
