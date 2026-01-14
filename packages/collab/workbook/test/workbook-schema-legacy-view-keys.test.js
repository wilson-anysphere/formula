import test from "node:test";
import assert from "node:assert/strict";
import * as Y from "yjs";

import { ensureWorkbookSchema, getWorkbookRoots } from "../src/index.ts";

test("ensureWorkbookSchema: preserves legacy top-level background_image keys when merging duplicate sheets", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    loser.set("background_image", "bg.png");

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);
  assert.equal(sheet.get("background_image"), "bg.png");

  doc.destroy();
});

test("ensureWorkbookSchema: preserves legacy top-level merged_ranges keys when merging duplicate sheets", () => {
  const doc = new Y.Doc();
  const { sheets } = getWorkbookRoots(doc);

  const merged = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }];

  doc.transact(() => {
    const loser = new Y.Map();
    loser.set("id", "Sheet1");
    loser.set("name", "Sheet1");
    loser.set("merged_ranges", merged);

    const winner = new Y.Map();
    winner.set("id", "Sheet1");
    winner.set("name", "Sheet1");

    sheets.push([loser]);
    sheets.push([winner]);
  });

  ensureWorkbookSchema(doc);

  assert.equal(sheets.length, 1);
  const sheet = sheets.get(0);
  assert.ok(sheet);
  assert.deepEqual(sheet.get("merged_ranges"), merged);

  doc.destroy();
});

