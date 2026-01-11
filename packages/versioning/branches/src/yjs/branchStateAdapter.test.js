import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { applyBranchStateToYjsDoc, branchStateFromYjsDoc } from "./branchStateAdapter.js";

test("branchStateFromYjsDoc: reads clobbered legacy comments array stored on a Map root", () => {
  const source = new Y.Doc();
  const commentsArray = source.getArray("comments");

  const yComment = new Y.Map();
  yComment.set("id", "c1");
  yComment.set("cellRef", "A1");
  yComment.set("content", "hello");
  yComment.set("resolved", false);
  yComment.set("mentions", []);
  yComment.set("replies", new Y.Array());
  commentsArray.push([yComment]);

  const update = Y.encodeStateAsUpdate(source);

  const doc = new Y.Doc();
  // Clobber the schema by instantiating the root as a Map before applying an
  // Array-backed update (older docs). This reproduces the real-world edge case
  // where the list items still exist but are invisible via `map.keys()`.
  doc.getMap("comments");
  Y.applyUpdate(doc, update);

  const state = branchStateFromYjsDoc(doc);
  assert.equal(state.comments.c1?.content, "hello");
});

test("applyBranchStateToYjsDoc: writes comments as Y.Maps for CommentManager compatibility", () => {
  const doc = new Y.Doc();

  applyBranchStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } },
    },
    cells: { Sheet1: {} },
    namedRanges: {},
    comments: {
      c1: { id: "c1", cellRef: "A1", content: "hello", resolved: false, replies: [] },
    },
  });

  const commentsMap = doc.getMap("comments");
  const value = commentsMap.get("c1");
  assert.ok(value instanceof Y.Map);
  assert.equal(value.get("content"), "hello");
  assert.ok(value.get("replies") instanceof Y.Array);
});

test("applyBranchStateToYjsDoc: preserves unknown sheet + cell metadata while applying snapshot", () => {
  const doc = new Y.Doc();

  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "OldName");
    sheet.set("color", "red");
    sheets.push([sheet]);

    const cells = doc.getMap("cells");
    const cell = new Y.Map();
    cell.set("value", 1);
    cell.set("modified", 123);
    cells.set("Sheet1:0:0", cell);
  });

  applyBranchStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "NewName" } },
    },
    cells: { Sheet1: { A1: { value: 2 } } },
    namedRanges: {},
    comments: {},
  });

  const sheet1 = doc.getArray("sheets").get(0);
  assert.ok(sheet1 instanceof Y.Map);
  assert.equal(sheet1.get("name"), "NewName");
  assert.equal(sheet1.get("color"), "red");

  const cell = doc.getMap("cells").get("Sheet1:0:0");
  assert.ok(cell instanceof Y.Map);
  assert.equal(cell.get("value"), 2);
  assert.equal(cell.get("modified"), 123);
});

test("applyBranchStateToYjsDoc: ensures at least one sheet when applying an empty branch state", () => {
  const doc = new Y.Doc();

  applyBranchStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: { order: [], metaById: {} },
    cells: {},
    namedRanges: {},
    comments: {},
  });

  const sheets = doc.getArray("sheets");
  assert.equal(sheets.length, 1);
  assert.equal(sheets.get(0)?.get("id"), "Sheet1");
});
