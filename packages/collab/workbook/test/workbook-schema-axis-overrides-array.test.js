import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { ensureWorkbookSchema, getWorkbookRoots } from "../src/index.ts";

test("ensureWorkbookSchema: preserves array-encoded top-level axis overrides when merging duplicate sheets", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    loser.set("colWidths", { "1": 111 });

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");
    // Winner uses the legacy/alternate array encoding.
    winner.set("colWidths", [
      [1, 100],
      [2, 200],
    ]);

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);

  // The winner's existing (non-empty) axis overrides should not be overwritten by the loser.
  assert.deepEqual(sheet.get("colWidths"), [
    [1, 100],
    [2, 200],
  ]);

  doc.destroy();
});

test("ensureWorkbookSchema: preserves array-encoded axis overrides inside `view` when merging duplicate sheets", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    const loserView = new Y.Map();
    loserView.set("colWidths", { "1": 111 });
    loser.set("view", loserView);

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");
    const winnerView = new Y.Map();
    // Winner uses the legacy/alternate array encoding.
    winnerView.set("colWidths", [
      [1, 100],
      [2, 200],
    ]);
    winner.set("view", winnerView);

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);

  const view = sheet.get("view");
  assert.ok(view instanceof Y.Map);

  assert.deepEqual(view.get("colWidths"), [
    [1, 100],
    [2, 200],
  ]);

  doc.destroy();
});

test("ensureWorkbookSchema: preserves array-encoded axis overrides inside plain-object `view` when merging duplicate sheets", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    loser.set("view", { colWidths: { "1": 111 } });

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");
    // Winner uses plain-object view with the legacy/alternate array encoding.
    winner.set("view", {
      colWidths: [
        [1, 100],
        [2, 200],
      ],
    });

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);

  const view = sheet.get("view");
  assert.ok(view && typeof view === "object" && !(view instanceof Y.AbstractType));

  assert.deepEqual(view.colWidths, [
    [1, 100],
    [2, 200],
  ]);

  doc.destroy();
});
