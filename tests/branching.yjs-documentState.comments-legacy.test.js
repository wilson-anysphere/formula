import test from "node:test";
import assert from "node:assert/strict";

import * as Y from "yjs";

import { applyDocumentStateToYjsDoc, yjsDocToDocumentState } from "../packages/versioning/branches/src/browser.js";

test("yjs documentState: snapshots legacy array-backed comments without clobbering root type", () => {
  const doc = new Y.Doc();
  const comments = doc.getArray("comments"); // legacy schema root

  const yComment = new Y.Map();
  yComment.set("id", "c1");
  yComment.set("cellRef", "A1");
  yComment.set("content", "hello");
  yComment.set("replies", new Y.Array());
  comments.push([yComment]);

  const state = yjsDocToDocumentState(doc);
  assert.equal(state.schemaVersion, 1);
  assert.ok(state.comments.c1);

  // Root type should remain an array (no accidental `getMap("comments")`).
  assert.ok(doc.share.get("comments") instanceof Y.Array);
});

test("yjs documentState: applyDocumentStateToYjsDoc preserves legacy comments root type", () => {
  const doc = new Y.Doc();
  doc.getArray("comments"); // legacy schema root

  applyDocumentStateToYjsDoc(doc, {
    schemaVersion: 1,
    sheets: { order: [], metaById: {} },
    cells: {},
    metadata: {},
    namedRanges: {},
    comments: {
      c1: { id: "c1", cellRef: "A1", content: "from snapshot", replies: [] },
    },
  });

  assert.ok(doc.share.get("comments") instanceof Y.Array);
});
