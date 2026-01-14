import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";
import { requireYjsCjs } from "../../../../collab/yjs-utils/test/require-yjs-cjs.js";

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

test("applyBranchStateToYjsDoc: clears removed cells via formula/value null markers (no root deletes)", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);

    const cells = doc.getMap("cells");
    const cell = new Y.Map();
    cell.set("formula", "=1");
    cell.set("value", null);
    cells.set("Sheet1:0:0", cell);
  });

  applyBranchStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1" } },
    },
    // Snapshot with no cells: should clear (not delete) the existing entry.
    cells: { Sheet1: {} },
    namedRanges: {},
    comments: {},
  });

  const cell2 = doc.getMap("cells").get("Sheet1:0:0");
  assert.ok(cell2 instanceof Y.Map);
  assert.equal(cell2.has("formula"), true);
  assert.equal(cell2.get("formula"), null);
  assert.equal(cell2.get("value"), null);
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

test("branchStateFromYjsDoc/applyBranchStateToYjsDoc: round-trips sheet view (frozen panes + axis sizes)", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("view", {
      frozenRows: 2,
      frozenCols: 1,
      backgroundImageId: "bg.png",
      colWidths: { "0": 120 },
      rowHeights: { "1": 40 },
    });
    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets.metaById.Sheet1?.view, {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
    mergedRanges: [],
    drawings: [],
  });

  const doc2 = new Y.Doc();
  applyBranchStateToYjsDoc(doc2, state);

  const sheet2 = doc2.getArray("sheets").get(0);
  assert.ok(sheet2 instanceof Y.Map);
  assert.deepEqual(sheet2.get("view"), {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
});

test("branchStateFromYjsDoc/applyBranchStateToYjsDoc: round-trips mergedRanges + drawings in sheet view", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("view", {
      mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
      drawings: [
        {
          id: 1,
          zOrder: 0,
          kind: { type: "image", imageId: "img-1" },
          anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
        },
      ],
    });
    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets.metaById.Sheet1?.view, {
    frozenRows: 0,
    frozenCols: 0,
    backgroundImageId: null,
    mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
    drawings: [
      {
        id: 1,
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ],
  });

  const doc2 = new Y.Doc();
  applyBranchStateToYjsDoc(doc2, state);
  const sheet2 = doc2.getArray("sheets").get(0);
  assert.ok(sheet2 instanceof Y.Map);
  assert.deepEqual(sheet2.get("view"), {
    frozenRows: 0,
    frozenCols: 0,
    backgroundImageId: null,
    mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }],
    drawings: [
      {
        id: 1,
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ],
  });
});

test("branchStateFromYjsDoc: reads legacy top-level sheet view fields (frozen panes + axis sizes)", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("frozenRows", 2);
    sheet.set("frozenCols", 1);
    sheet.set("background_image_id", "bg.png");
    sheet.set("colWidths", { "0": 120 });
    sheet.set("rowHeights", { "1": 40 });
    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets.metaById.Sheet1?.view, {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
    mergedRanges: [],
    drawings: [],
  });

  const doc2 = new Y.Doc();
  applyBranchStateToYjsDoc(doc2, state);
  const sheet2 = doc2.getArray("sheets").get(0);
  assert.ok(sheet2 instanceof Y.Map);
  // Canonical write format is nested under `view`.
  assert.deepEqual(sheet2.get("view"), {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
  assert.equal(sheet2.get("frozenRows"), undefined);
  assert.equal(sheet2.get("frozenCols"), undefined);
  assert.equal(sheet2.get("backgroundImageId"), undefined);
  assert.equal(sheet2.get("background_image_id"), undefined);
});

test("applyBranchStateToYjsDoc: drops legacy top-level axis sizes when applying snapshot", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("color", "red");
    // Legacy top-level fields that should be canonicalized into `view`.
    sheet.set("frozenRows", 2);
    sheet.set("frozenCols", 1);
    sheet.set("backgroundImageId", "bg.png");
    sheet.set("colWidths", { "0": 120 });
    sheet.set("rowHeights", { "1": 40 });
    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets.metaById.Sheet1?.view, {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
    mergedRanges: [],
    drawings: [],
  });

  applyBranchStateToYjsDoc(doc, state);

  const sheet1 = doc.getArray("sheets").get(0);
  assert.ok(sheet1 instanceof Y.Map);
  assert.equal(sheet1.get("color"), "red");
  assert.deepEqual(sheet1.get("view"), {
    frozenRows: 2,
    frozenCols: 1,
    backgroundImageId: "bg.png",
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
  assert.equal(sheet1.get("frozenRows"), undefined);
  assert.equal(sheet1.get("frozenCols"), undefined);
  assert.equal(sheet1.get("backgroundImageId"), undefined);
  assert.equal(sheet1.get("background_image_id"), undefined);
  assert.equal(sheet1.get("colWidths"), undefined);
  assert.equal(sheet1.get("rowHeights"), undefined);
});

test("branchStateFromYjsDoc: reads layered format defaults from top-level sheet metadata (defaultFormat/rowFormats/colFormats/formatRunsByCol)", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    sheet.set("defaultFormat", { font: { bold: true } });

    const rowFormats = new Y.Map();
    rowFormats.set("0", { font: { italic: true } });
    sheet.set("rowFormats", rowFormats);

    const colFormats = new Y.Map();
    colFormats.set("0", { fill: { fgColor: "red" } });
    sheet.set("colFormats", colFormats);

    const runsByCol = new Y.Map();
    runsByCol.set("0", [{ startRow: 0, endRowExclusive: 1, format: { numberFormat: "0%" } }]);
    sheet.set("formatRunsByCol", runsByCol);

    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets.metaById.Sheet1?.view?.defaultFormat, { font: { bold: true } });
  assert.deepEqual(state.sheets.metaById.Sheet1?.view?.rowFormats, { "0": { font: { italic: true } } });
  assert.deepEqual(state.sheets.metaById.Sheet1?.view?.colFormats, { "0": { fill: { fgColor: "red" } } });
  assert.deepEqual(state.sheets.metaById.Sheet1?.view?.formatRunsByCol, [
    { col: 0, runs: [{ startRow: 0, endRowExclusive: 1, format: { numberFormat: "0%" } }] },
  ]);
});

test("applyBranchStateToYjsDoc: writes layered format defaults to top-level sheet metadata", () => {
  const doc = new Y.Doc();

  // Seed an entry with stale formatting so the apply path must overwrite/clear it.
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("defaultFormat", { font: { bold: false } });
    const rowFormats = new Y.Map();
    rowFormats.set("0", { font: { italic: false } });
    sheet.set("rowFormats", rowFormats);
    sheets.push([sheet]);
  });

  applyBranchStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            defaultFormat: { font: { bold: true } },
            rowFormats: { "0": { font: { italic: true } } },
            colFormats: { "0": { fill: { fgColor: "red" } } },
            formatRunsByCol: [{ col: 0, runs: [{ startRow: 0, endRowExclusive: 1, format: { numberFormat: "0%" } }] }],
          },
        },
      },
    },
    cells: { Sheet1: {} },
    namedRanges: {},
    comments: {},
  });

  const sheet1 = doc.getArray("sheets").get(0);
  assert.ok(sheet1 instanceof Y.Map);

  assert.deepEqual(sheet1.get("defaultFormat"), { font: { bold: true } });

  const rowFormats = sheet1.get("rowFormats");
  assert.ok(rowFormats instanceof Y.Map);
  assert.deepEqual(rowFormats.get("0"), { font: { italic: true } });

  const colFormats = sheet1.get("colFormats");
  assert.ok(colFormats instanceof Y.Map);
  assert.deepEqual(colFormats.get("0"), { fill: { fgColor: "red" } });

  const runsByCol = sheet1.get("formatRunsByCol");
  assert.ok(runsByCol instanceof Y.Map);
  assert.deepEqual(runsByCol.get("0"), [{ startRow: 0, endRowExclusive: 1, format: { numberFormat: "0%" } }]);
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

test("branchStateFromYjsDoc/applyBranchStateToYjsDoc: round-trips sheet visibility + tabColor", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("visibility", "hidden");
    // Use lowercase to ensure normalization canonicalizes.
    sheet.set("tabColor", "ff00ff00");
    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.equal(state.sheets.metaById.Sheet1?.visibility, "hidden");
  assert.equal(state.sheets.metaById.Sheet1?.tabColor, "FF00FF00");

  const doc2 = new Y.Doc();
  applyBranchStateToYjsDoc(doc2, state);
  const sheet2 = doc2.getArray("sheets").get(0);
  assert.ok(sheet2 instanceof Y.Map);
  assert.equal(sheet2.get("visibility"), "hidden");
  assert.equal(sheet2.get("tabColor"), "FF00FF00");
});

test("applyBranchStateToYjsDoc: clears sheet tabColor when meta.tabColor is null", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheet.set("tabColor", "FFFF0000");
    sheets.push([sheet]);
  });

  applyBranchStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: { Sheet1: { id: "Sheet1", name: "Sheet1", tabColor: null } },
    },
    cells: { Sheet1: {} },
    namedRanges: {},
    comments: {},
  });

  const sheet1 = doc.getArray("sheets").get(0);
  assert.ok(sheet1 instanceof Y.Map);
  assert.equal(sheet1.get("tabColor"), undefined);
});

test("branch adapter works when workbook roots were instantiated by a different Yjs instance (CJS getArray/getMap)", () => {
  const Ycjs = requireYjsCjs();

  const remote = new Ycjs.Doc();
  remote.transact(() => {
    const sheets = remote.getArray("sheets");
    const sheet = new Ycjs.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");
    sheets.push([sheet]);

    const cells = remote.getMap("cells");
    const cell = new Ycjs.Map();
    cell.set("value", 42);
    cells.set("Sheet1:0:0", cell);
  });

  const update = Ycjs.encodeStateAsUpdate(remote);

  /**
   * @returns {Y.Doc}
   */
  function createDocWithForeignRoots() {
    const doc = new Y.Doc();
    // Simulate a mixed-module environment where updates are applied by another
    // Yjs instance (e.g. y-websocket using CJS `require("yjs")`).
    Ycjs.applyUpdate(doc, update);
    // Simulate another instance eagerly instantiating workbook roots before our
    // code touches them, resulting in foreign root constructors.
    Ycjs.Doc.prototype.getArray.call(doc, "sheets");
    Ycjs.Doc.prototype.getMap.call(doc, "cells");
    return doc;
  }

  const doc1 = createDocWithForeignRoots();
  assert.ok(!(doc1.share.get("sheets") instanceof Y.Array));
  assert.ok(!(doc1.share.get("cells") instanceof Y.Map));
  assert.throws(() => doc1.getArray("sheets"), /different constructor/);
  assert.throws(() => doc1.getMap("cells"), /different constructor/);

  const state = branchStateFromYjsDoc(doc1);
  assert.deepEqual(state.sheets.order, ["Sheet1"]);
  assert.equal(state.cells.Sheet1?.A1?.value, 42);

  // Reading should normalize roots to the local Yjs instance so subsequent
  // `doc.getArray/getMap` calls don't throw.
  assert.ok(doc1.share.get("sheets") instanceof Y.Array);
  assert.ok(doc1.share.get("cells") instanceof Y.Map);
  assert.ok(doc1.getArray("sheets") instanceof Y.Array);
  assert.ok(doc1.getMap("cells") instanceof Y.Map);

  const doc2 = createDocWithForeignRoots();
  applyBranchStateToYjsDoc(doc2, state, { origin: "test" });
  assert.ok(doc2.share.get("sheets") instanceof Y.Array);
  assert.ok(doc2.share.get("cells") instanceof Y.Map);
  assert.equal(doc2.getMap("cells").get("Sheet1:0:0")?.get("value"), 42);

  remote.destroy();
  doc1.destroy();
  doc2.destroy();
});

test("branchStateFromYjsDoc: drops drawings with oversized Y.Text ids without materializing strings", () => {
  const doc = new Y.Doc();
  doc.transact(() => {
    const sheets = doc.getArray("sheets");
    const sheet = new Y.Map();
    sheet.set("id", "Sheet1");
    sheet.set("name", "Sheet1");

    const view = new Y.Map();
    const drawings = new Y.Array();
    const drawing = new Y.Map();
    const idText = new Y.Text();
    idText.insert(0, "x".repeat(5000));
    // If snapshot extraction calls `toString()` on this oversized id, this test should fail.
    idText.toString = () => {
      throw new Error("unexpected Y.Text.toString() on oversized drawing id");
    };
    drawing.set("id", idText);
    drawing.set("zOrder", 0);
    drawings.push([drawing]);
    view.set("drawings", drawings);
    sheet.set("view", view);
    sheets.push([sheet]);
  });

  const state = branchStateFromYjsDoc(doc);
  assert.deepEqual(state.sheets.metaById.Sheet1?.view?.drawings, []);
});
