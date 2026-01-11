import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { diffYjsWorkbookSnapshots } from "../packages/versioning/src/index.js";

test("diffYjsWorkbookSnapshots reports workbook-level metadata changes", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");
  const comments = doc.getMap("comments");
  const namedRanges = doc.getMap("namedRanges");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  const sheet2 = new Y.Map();
  sheet2.set("id", "sheet2");
  sheet2.set("name", "Sheet2");
  sheets.push([sheet1, sheet2]);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "move-me");
    cell.set("formula", "=A1+B1");
    cells.set("sheet1:0:0", cell);

    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "A1");
    comment.set("content", "Original comment");
    comments.set("c1", comment);

    namedRanges.set("NR1", { sheetId: "sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } });
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    // Sheet rename + add/remove.
    sheet1.set("name", "Renamed");
    sheets.delete(1, 1); // remove sheet2
    const sheet3 = new Y.Map();
    sheet3.set("id", "sheet3");
    sheet3.set("name", "Sheet3");
    sheets.push([sheet3]);

    // Move a cell via cut/paste with a semantically equivalent formula.
    cells.delete("sheet1:0:0");
    const moved = new Y.Map();
    moved.set("value", "move-me");
    moved.set("formula", "=B1 + A1");
    cells.set("sheet1:2:3", moved);

    // Edit a comment (content + resolution + replies).
    const comment = comments.get("c1");
    assert.ok(comment instanceof Y.Map);
    comment.set("content", "Updated comment");
    comment.set("resolved", true);
    const replies = new Y.Array();
    replies.push([{ id: "r1", content: "First reply" }]);
    comment.set("replies", replies);

    // Named ranges.
    namedRanges.set("NR2", { sheetId: "sheet1", rect: { r0: 1, c0: 1, r1: 2, c1: 2 } });
    namedRanges.set("NR1", { sheetId: "sheet1", rect: { r0: 0, c0: 0, r1: 3, c1: 3 } });
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });

  assert.deepEqual(diff.sheets.renamed, [{ id: "sheet1", beforeName: "Sheet1", afterName: "Renamed" }]);
  assert.deepEqual(diff.sheets.added, [{ id: "sheet3", name: "Sheet3" }]);
  assert.deepEqual(diff.sheets.removed, [{ id: "sheet2", name: "Sheet2" }]);

  assert.deepEqual(
    diff.cellsBySheet.map((entry) => entry.sheetId),
    ["sheet1", "sheet2", "sheet3"],
  );
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.moved.length, 1);
  assert.deepEqual(sheet1Diff.moved[0].oldLocation, { row: 0, col: 0 });
  assert.deepEqual(sheet1Diff.moved[0].newLocation, { row: 2, col: 3 });

  assert.deepEqual(diff.comments.added, []);
  assert.deepEqual(diff.comments.removed, []);
  assert.equal(diff.comments.modified.length, 1);
  assert.equal(diff.comments.modified[0].id, "c1");
  assert.deepEqual(diff.comments.modified[0].before, {
    id: "c1",
    cellRef: "A1",
    content: "Original comment",
    resolved: false,
    repliesLength: 0,
  });
  assert.deepEqual(diff.comments.modified[0].after, {
    id: "c1",
    cellRef: "A1",
    content: "Updated comment",
    resolved: true,
    repliesLength: 1,
  });

  assert.deepEqual(diff.namedRanges.added.map((r) => r.key), ["NR2"]);
  assert.deepEqual(diff.namedRanges.removed, []);
  assert.equal(diff.namedRanges.modified.length, 1);
  assert.equal(diff.namedRanges.modified[0].key, "NR1");
});

