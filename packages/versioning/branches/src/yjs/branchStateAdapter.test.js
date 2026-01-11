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

test("branchStateFromYjsDoc: reads map entries stored on an Array root (mixed schema)", () => {
  const source = new Y.Doc();
  const commentsMap = source.getMap("comments");
  const comment = new Y.Map();
  comment.set("id", "c1");
  comment.set("cellRef", "A1");
  comment.set("content", "hello");
  comment.set("resolved", false);
  comment.set("mentions", []);
  comment.set("replies", new Y.Array());
  commentsMap.set("c1", comment);
  const update = Y.encodeStateAsUpdate(source);

  const doc = new Y.Doc();
  doc.getArray("comments");
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

test("branchStateFromYjsDoc/applyBranchStateToYjsDoc: preserves encrypted cell payloads", () => {
  const enc = {
    v: 1,
    alg: "AES-256-GCM",
    keyId: "k1",
    ivBase64: "iv",
    tagBase64: "tag",
    ciphertextBase64: "ct",
  };

  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);

    const cells = doc.getMap("cells");
    const cell = new Y.Map();
    cell.set("enc", enc);
    cells.set("Sheet1:0:0", cell);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.cells.Sheet1.A1, { enc });

  const doc2 = new Y.Doc();
  applyBranchStateToYjsDoc(doc2, state);

  const cell2 = doc2.getMap("cells").get("Sheet1:0:0");
  assert.ok(cell2 instanceof Y.Map);
  assert.deepEqual(cell2.get("enc"), enc);
  assert.equal(cell2.get("value"), undefined);
  assert.equal(cell2.get("formula"), undefined);
});

test("branchStateFromYjsDoc: prefers encrypted payloads over plaintext duplicates across legacy cell keys", () => {
  const enc = {
    v: 1,
    alg: "AES-256-GCM",
    keyId: "k1",
    ivBase64: "iv",
    tagBase64: "tag",
    ciphertextBase64: "ct",
  };

  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);

    const cells = doc.getMap("cells");
    const cellEnc = new Y.Map();
    cellEnc.set("enc", enc);
    // Encrypted payload stored under a legacy key encoding.
    cells.set("Sheet1:0,0", cellEnc);

    const cellPlain = new Y.Map();
    cellPlain.set("value", "leaked");
    // Plaintext duplicate stored under the canonical key.
    cells.set("Sheet1:0:0", cellPlain);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.cells.Sheet1.A1, { enc });
});

test("branchStateFromYjsDoc/applyBranchStateToYjsDoc: round-trips workbook metadata", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);

    const metadata = doc.getMap("metadata");
    metadata.set("title", "Budget");
    metadata.set("theme", { name: "dark" });
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.metadata, { theme: { name: "dark" }, title: "Budget" });

  const doc2 = new Y.Doc();
  applyBranchStateToYjsDoc(doc2, state);

  assert.equal(doc2.getMap("metadata").get("title"), "Budget");
  assert.deepEqual(doc2.getMap("metadata").get("theme"), { name: "dark" });
});
