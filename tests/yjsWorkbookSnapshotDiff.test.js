import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { diffYjsWorkbookSnapshots } from "../packages/versioning/src/index.js";
import { workbookStateFromYjsDoc } from "../packages/versioning/src/yjs/workbookState.js";

test("diffYjsWorkbookSnapshots reports workbook-level metadata changes", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");
  const comments = doc.getMap("comments");
  const metadata = doc.getMap("metadata");
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

    metadata.set("title", "Budget");
    metadata.set("owner", "u1");
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

    // Workbook metadata.
    metadata.set("title", "Budget (edited)");
    metadata.delete("owner");
    metadata.set("theme", { name: "dark" });
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });

  assert.deepEqual(diff.sheets.renamed, [{ id: "sheet1", beforeName: "Sheet1", afterName: "Renamed" }]);
  assert.deepEqual(diff.sheets.added, [{ id: "sheet3", name: "Sheet3", afterIndex: 1 }]);
  assert.deepEqual(diff.sheets.removed, [{ id: "sheet2", name: "Sheet2", beforeIndex: 1 }]);
  assert.deepEqual(diff.sheets.moved, []);

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

  assert.deepEqual(diff.metadata.added.map((r) => r.key), ["theme"]);
  assert.deepEqual(diff.metadata.removed.map((r) => r.key), ["owner"]);
  assert.equal(diff.metadata.modified.length, 1);
  assert.equal(diff.metadata.modified[0].key, "title");
});

test("diffYjsWorkbookSnapshots reports sheet metadata changes (visibility/tabColor/frozen panes)", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  sheet1.set("visibility", "visible");
  sheet1.set("tabColor", "FF00FF00");
  sheet1.set("view", { frozenRows: 1, frozenCols: 0 });
  sheets.push([sheet1]);

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    sheet1.set("visibility", "hidden");
    sheet1.set("tabColor", null);
    sheet1.set("view", { frozenRows: 2, frozenCols: 3 });
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.added, []);
  assert.deepEqual(diff.sheets.removed, []);
  assert.deepEqual(diff.sheets.renamed, []);
  assert.deepEqual(diff.sheets.moved, []);
  assert.deepEqual(diff.sheets.metaChanged, [
    { id: "sheet1", field: "tabColor", before: "FF00FF00", after: null },
    { id: "sheet1", field: "view.frozenCols", before: 0, after: 3 },
    { id: "sheet1", field: "view.frozenRows", before: 1, after: 2 },
    { id: "sheet1", field: "visibility", before: "visible", after: "hidden" },
  ]);
});

test("diffYjsWorkbookSnapshots reads frozen panes from legacy top-level sheet keys", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  // Legacy/experimental: frozen panes stored on the sheet entry (not nested in `view`).
  sheet1.set("frozenRows", 0);
  sheet1.set("frozenCols", 1);
  sheets.push([sheet1]);

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    sheet1.set("frozenRows", 4);
    sheet1.set("frozenCols", 2);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.metaChanged, [
    { id: "sheet1", field: "view.frozenCols", before: 1, after: 2 },
    { id: "sheet1", field: "view.frozenRows", before: 0, after: 4 },
  ]);
});

test("diffYjsWorkbookSnapshots canonicalizes tabColor to 8-digit ARGB", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  // Non-canonical 6-digit RGB hex (with "#"); should normalize to ARGB with opaque alpha.
  sheet1.set("tabColor", "#00FF00");
  sheets.push([sheet1]);

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    sheet1.set("tabColor", null);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.metaChanged, [{ id: "sheet1", field: "tabColor", before: "FF00FF00", after: null }]);
});

test("diffYjsWorkbookSnapshots accepts tabColor.argb and canonicalizes to 8-digit ARGB", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  // Some snapshot producers use `{ argb: "..." }` objects (ExcelJS-style).
  sheet1.set("tabColor", { argb: "#00FF00" });
  sheets.push([sheet1]);

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    sheet1.set("tabColor", null);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  assert.deepEqual(diff.sheets.metaChanged, [{ id: "sheet1", field: "tabColor", before: "FF00FF00", after: null }]);
});

test("diffYjsWorkbookSnapshots reports formatOnly edits when column default formats change (layered formats)", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  sheets.push([sheet1]);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    cells.set("sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const colFormats = new Y.Map();
    colFormats.set("0", { font: { bold: true } });
    sheet1.set("colFormats", colFormats);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffYjsWorkbookSnapshots reports formatOnly edits when row default formats change (layered formats)", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  sheets.push([sheet1]);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    cells.set("sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    const rowFormats = new Y.Map();
    rowFormats.set("0", { font: { italic: true } });
    sheet1.set("rowFormats", rowFormats);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffYjsWorkbookSnapshots reports formatOnly edits when sheet default formats change (layered formats)", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  sheets.push([sheet1]);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    cells.set("sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    sheet1.set("defaultFormat", { fill: { fgColor: "#FFFF0000" } });
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffYjsWorkbookSnapshots reports formatOnly edits when range-run formats change (formatRunsByCol)", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");

  const sheet1 = new Y.Map();
  sheet1.set("id", "sheet1");
  sheet1.set("name", "Sheet1");
  sheets.push([sheet1]);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "x");
    cells.set("sheet1:0:0", cell);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    // Apply a compressed format run (Column A, row 0 only).
    const runsByCol = new Y.Map();
    runsByCol.set("0", [{ startRow: 0, endRowExclusive: 1, format: { font: { bold: true } } }]);
    sheet1.set("formatRunsByCol", runsByCol);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);

  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });
  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.formatOnly.length, 1);
  assert.deepEqual(sheet1Diff.formatOnly[0].cell, { row: 0, col: 0 });
});

test("diffYjsWorkbookSnapshots supports comments stored as a Y.Array", () => {
  const doc = new Y.Doc();
  const sheets = doc.getArray("sheets");
  const cells = doc.getMap("cells");
  const comments = doc.getArray("comments");

  const sheet = new Y.Map();
  sheet.set("id", "sheet1");
  sheet.set("name", "Sheet1");
  sheets.push([sheet]);

  doc.transact(() => {
    const cell = new Y.Map();
    cell.set("value", "move-me");
    cell.set("formula", "=A1+B1");
    cells.set("sheet1:0:0", cell);

    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "A1");
    comment.set("content", "Original comment");
    comment.set("resolved", false);
    comment.set("replies", new Y.Array());
    comments.push([comment]);
  });

  const beforeSnapshot = Y.encodeStateAsUpdate(doc);

  doc.transact(() => {
    cells.delete("sheet1:0:0");
    const moved = new Y.Map();
    moved.set("value", "move-me");
    moved.set("formula", "=B1 + A1");
    cells.set("sheet1:2:3", moved);

    const comment = comments.get(0);
    assert.ok(comment instanceof Y.Map);
    comment.set("content", "Updated comment");
    comment.set("resolved", true);
    const replies = comment.get("replies");
    assert.ok(replies instanceof Y.Array);
    replies.push([{ id: "r1", content: "First reply" }]);
  });

  const afterSnapshot = Y.encodeStateAsUpdate(doc);
  const diff = diffYjsWorkbookSnapshots({ beforeSnapshot, afterSnapshot });

  assert.equal(diff.comments.modified.length, 1);
  assert.equal(diff.comments.modified[0].id, "c1");
  assert.equal(diff.comments.modified[0].after.repliesLength, 1);

  const sheet1Diff = diff.cellsBySheet.find((entry) => entry.sheetId === "sheet1")?.diff;
  assert.ok(sheet1Diff);
  assert.equal(sheet1Diff.moved.length, 1);
});

test("workbookStateFromYjsDoc supports legacy list comments stored inside a Y.Map root (clobbered schema)", () => {
  const source = new Y.Doc();
  const comments = source.getArray("comments");

  const comment = new Y.Map();
  comment.set("id", "c1");
  comment.set("cellRef", "A1");
  comment.set("content", "Original comment");
  comment.set("resolved", false);
  comment.set("replies", new Y.Array());
  comments.push([comment]);

  const snapshot = Y.encodeStateAsUpdate(source);

  const doc = new Y.Doc();
  Y.applyUpdate(doc, snapshot);

  // Simulate the historical bug: instantiating as a map first makes the array
  // inaccessible via `doc.getArray("comments")`, but the list items still exist.
  doc.getMap("comments");

  const state = workbookStateFromYjsDoc(doc);
  assert.equal(state.comments.size, 1);
  assert.deepEqual(state.comments.get("c1"), {
    id: "c1",
    cellRef: "A1",
    content: "Original comment",
    resolved: false,
    repliesLength: 0,
  });
});

test("workbookStateFromYjsDoc supports map entries stored inside a Y.Array root (mixed schema)", () => {
  const source = new Y.Doc();
  const commentsMap = source.getMap("comments");
  const comment = new Y.Map();
  comment.set("id", "c1");
  comment.set("cellRef", "A1");
  comment.set("content", "From map peer");
  comment.set("resolved", false);
  comment.set("replies", new Y.Array());
  commentsMap.set("c1", comment);
  const update = Y.encodeStateAsUpdate(source);

  const doc = new Y.Doc();
  doc.getArray("comments"); // force array constructor
  Y.applyUpdate(doc, update);

  const state = workbookStateFromYjsDoc(doc);
  assert.equal(state.comments.size, 1);
  assert.equal(state.comments.get("c1")?.content, "From map peer");
});
