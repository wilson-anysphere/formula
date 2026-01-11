import assert from "node:assert/strict";
import test from "node:test";

import * as Y from "yjs";

import { applyDocumentStateToYjsDoc, yjsDocToDocumentState } from "./documentState.js";

test("yjs document adapter: normalizes formulas + handles legacy cell keys", () => {
  const doc = new Y.Doc();
  const cells = doc.getMap("cells");

  const cellA1 = new Y.Map();
  cellA1.set("value", 123);
  cells.set("Sheet1:0:0", cellA1);

  // Legacy `${sheetId}:${row},${col}` encoding + formula missing "=".
  const cellB1 = new Y.Map();
  cellB1.set("formula", "1+1");
  cellB1.set("format", { bold: true });
  cells.set("Sheet1:0,1", cellB1);

  // Unit-test convenience encoding.
  const cellC1 = new Y.Map();
  cellC1.set("formula", "=SUM(1,2)");
  cells.set("r0c2", cellC1);

  const state = yjsDocToDocumentState(doc);
  assert.deepEqual(state, {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: { id: "Sheet1", name: "Sheet1" },
      },
    },
    cells: {
      Sheet1: {
        A1: { value: 123 },
        B1: { formula: "=1+1", format: { bold: true } },
        C1: { formula: "=SUM(1,2)" },
      },
    },
    namedRanges: {},
    comments: {},
  });

  const doc2 = new Y.Doc();
  applyDocumentStateToYjsDoc(doc2, state, { origin: { test: true } });

  const cells2 = doc2.getMap("cells");
  assert.equal(cells2.has("Sheet1:0:0"), true);
  assert.equal(cells2.has("Sheet1:0:1"), true);
  assert.equal(cells2.has("Sheet1:0:2"), true);
  assert.equal(cells2.has("Sheet1:0,1"), false);
  assert.equal(cells2.has("r0c2"), false);

  const b1 = /** @type {Y.Map<any>} */ (cells2.get("Sheet1:0:1"));
  assert.ok(b1);
  assert.equal(b1.get("formula"), "=1+1");
  assert.deepEqual(b1.get("format"), { bold: true });
});

test("yjs document adapter: reads comments from map, array, and clobbered roots", () => {
  {
    const doc = new Y.Doc();
    const comments = doc.getMap("comments");
    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "A1");
    comment.set("content", "Map comment");
    comments.set("c1", comment);

    const state = yjsDocToDocumentState(doc);
    assert.equal(state.comments.c1.content, "Map comment");
  }

  {
    const doc = new Y.Doc();
    const comments = doc.getArray("comments");
    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "A1");
    comment.set("content", "Array comment");
    comments.push([comment]);

    const state = yjsDocToDocumentState(doc);
    assert.equal(state.comments.c1.content, "Array comment");
  }

  {
    const legacy = new Y.Doc();
    const comments = legacy.getArray("comments");
    const comment = new Y.Map();
    comment.set("id", "c1");
    comment.set("cellRef", "A1");
    comment.set("content", "Legacy clobbered");
    comments.push([comment]);

    const snapshot = Y.encodeStateAsUpdate(legacy);
    const doc = new Y.Doc();
    Y.applyUpdate(doc, snapshot);

    // Simulate the historical bug: instantiate as a map first.
    doc.getMap("comments");

    const state = yjsDocToDocumentState(doc);
    assert.equal(state.comments.c1.content, "Legacy clobbered");
  }
});

test("yjs document adapter: applyDocumentStateToYjsDoc clears legacy list items on a clobbered comments map root", () => {
  const legacy = new Y.Doc();
  const legacyComments = legacy.getArray("comments");
  const legacyComment = new Y.Map();
  legacyComment.set("id", "c1");
  legacyComment.set("cellRef", "A1");
  legacyComment.set("content", "Legacy");
  legacyComments.push([legacyComment]);
  const legacySnapshot = Y.encodeStateAsUpdate(legacy);

  const target = new Y.Doc();
  Y.applyUpdate(target, legacySnapshot);
  target.getMap("comments"); // clobber

  applyDocumentStateToYjsDoc(
    target,
    {
      schemaVersion: 1,
      sheets: { order: [], metaById: {} },
      cells: {},
      namedRanges: {},
      comments: {
        c2: { id: "c2", cellRef: "A2", content: "Canonical" },
      },
    },
    { origin: { test: true } },
  );

  const restored = target.getMap("comments");
  assert.equal(restored.size, 1);
  assert.equal(restored.has("c1"), false);
  assert.equal(restored.has("c2"), true);

  // Ensure no legacy list items remain on the map root.
  let item = restored._start;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      assert.fail("expected restored comments map to have no legacy list items");
    }
    item = item.right;
  }
});
